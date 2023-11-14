use super::books::Book;
use crate::errors::Error;
use crate::types::TagID;
use calamine::{DataType, Range, Reader};
use std::collections::HashMap;
use std::path::PathBuf;

enum Column {
    Initial,
    Contract,
}
enum Row {
    TableHeader,
}

struct SearchTag {
    tag: TagID,
    is_required: bool,
    group_by_row: Option<Row>,
    group_by_col: Option<Column>,
}

// перечислены в порядке вхождения слева на право и сверху вниз на листе Excel (вход по строкам важен для валидации)
// группировка по строке и столбцу для валидации в будующих версиях программы (не реализовано)
#[rustfmt::skip]
const SEARCH_TAGS: [SearchTag; 10] = [
    SearchTag { is_required: false, group_by_row: None,                   group_by_col: Some(Column::Initial),  tag: TagID::Исполнитель },
    SearchTag { is_required: true,  group_by_row: None,                   group_by_col: Some(Column::Initial),  tag: TagID::Стройка },
    SearchTag { is_required: true,  group_by_row: None,                   group_by_col: Some(Column::Initial),  tag: TagID::Объект },
    SearchTag { is_required: true,  group_by_row: None,                   group_by_col: Some(Column::Contract), tag: TagID::ДоговорПодряда },
    SearchTag { is_required: true,  group_by_row: None,                   group_by_col: Some(Column::Contract), tag: TagID::ДопСоглашение },
    SearchTag { is_required: true,  group_by_row: None,                   group_by_col: None,                   tag: TagID::НомерДокумента },
    SearchTag { is_required: true,  group_by_row: Some(Row::TableHeader), group_by_col: None,                   tag: TagID::НаименованиеРаботИЗатрат },
    SearchTag { is_required: false, group_by_row: Some(Row::TableHeader), group_by_col: None,                   tag: TagID::ЗтрВсего },
    SearchTag { is_required: false, group_by_row: None,                   group_by_col: Some(Column::Initial),  tag: TagID::ИтогоПоАкту },
    SearchTag { is_required: true,  group_by_row: None,                   group_by_col: None,                   tag: TagID::СтоимостьМатериальныхРесурсовВсего },
];

pub struct Sheet {
    pub path: PathBuf,
    pub sheet_name: String,
    pub data: Range<DataType>,
    pub search_points: HashMap<TagID, (usize, usize)>,
    pub range_start: (usize, usize),
}

impl<'a> Sheet {
    pub fn new(
        // разработчики Calamine делают зачем-то &mut self в функции worksheet_range(&mut self, name: &str),
        // из-за этого workbook приходится держать мутабельным, хотя этот код его менять вовсе не собирается
        // (из-за мутабельности workbook проблема при попытке множественных ссылок: можно только клонировать)
        workbook: &'a mut Book,
        user_entered_sh_name: &'a str,
        expected_columns_sum: usize,
    ) -> Result<Sheet, Error<'a>> {
        let entered_sh_name_lowercase = user_entered_sh_name.to_lowercase();

        let sheet_name = workbook
            .data
            .sheet_names()
            .iter()
            .find(|name| name.to_lowercase() == entered_sh_name_lowercase)
            .ok_or(Error::CalamineSheetOfTheBookIsUndetectable {
                file_path: &workbook.path,
                sh_name_for_search: user_entered_sh_name,
                sh_names: workbook.data.sheet_names().to_owned(),
            })?
            .clone();

        let xl_sheet = workbook
            .data
            .worksheet_range(&sheet_name)
            .ok_or(Error::CalamineSheetOfTheBookIsUndetectable {
                file_path: &workbook.path,
                sh_name_for_search: user_entered_sh_name,
                sh_names: workbook.data.sheet_names().to_owned(),
            })?
            .or_else(|error| {
                Err(Error::CalamineSheetOfTheBookIsUnreadable {
                    file_path: &workbook.path,
                    sh_name: sheet_name.to_owned(),
                    err: error,
                })
            })?;

        // при ошибки передается точное имя листа с учетом регистра (не используем ввод пользователя)
        let sheet_start_coords = xl_sheet.start().ok_or(Error::EmptySheetRange {
            file_path: &workbook.path,
            sh_name: sheet_name.to_owned(),
        })?;

        let mut search_points = HashMap::new();

        let mut limited_cell_iterator = xl_sheet.used_cells();
        let mut found_cell;
        for item in SEARCH_TAGS {
            let mut non_limited_cell_iterator = xl_sheet.used_cells();

            // Для обязательных тегов расходуемый итератор обеспечит валидацию очередности вохождения тегов
            // (например, "Стройку" мы ожидаем выше "Объекта, не наоборот).
            // Необязательным тегам расходуемый итератор не подходит, т.к. необязательный тег при отсутсвии израсходует итератор
            let iterator = if item.is_required {
                &mut limited_cell_iterator
            } else {
                &mut non_limited_cell_iterator
            };

            found_cell = iterator.find(|cell| match cell.2.get_string() {
                Some(str) => {
                    //  println!("{}   {}    {}", str.eq_ignore_ascii_case(item.tag), str, item.tag);

                    str.to_lowercase() == item.tag.as_str()
                }
                None => false,
            });

            if let Some((row, col, _)) = found_cell {
                search_points.insert(item.tag, (row, col));
            }
        }

        // Валидация на полноту данных: выше итератор расходующий ячейки и если хоть один поиск провалился, то это преждевременно
        // потребит все ячейки и извлечение по тегу последней строки в SEARCH_TAGS гарантированно провалится
        let validation_tag = SEARCH_TAGS
            .iter()
            .filter(|search_tag| search_tag.is_required)
            .last()
            .ok_or_else(|| Error::InternalLogic {
                tech_descr: "SEARCH_TAGS пуст".to_string(),
                err: None,
            })?
            .tag;

        search_points
            .get(&validation_tag)
            .ok_or(Error::SheetNotContainAllNecessaryData {
                file_path: &workbook.path,
                search_points: search_points.clone(),
            })?;

        // Проверка значений на удаленность столбцов, чтобы гарантировать что найден нужный лист.
        let initial_column_coords = search_points
            .get(&TagID::Стройка)
            .unwrap_or_else(|| panic!("ложь: \"Необеспечены действительные имена HashMap\""));

        let (just_a_amount_requir_col, just_a_sum_requir_col) =
            SEARCH_TAGS
                .iter()
                .fold((0_usize, 0), |acc, item| match item.is_required {
                    true => (
                        acc.0 + 1,
                        acc.1
                            + search_points
                                .get(&item.tag)
                                .unwrap_or_else(|| {
                                    panic!("ложь: \"Необеспечены действительные имена HashMap\"")
                                })
                                .1,
                    ),
                    false => acc,
                });

        if let false = just_a_sum_requir_col - initial_column_coords.1 * just_a_amount_requir_col
            == expected_columns_sum
        {
            return Err(Error::ShiftedColumnsInHeader(&workbook.path));
        }

        let range_start = (sheet_start_coords.0 as usize, sheet_start_coords.1 as usize);

        Ok(Sheet {
            path: workbook.path.clone(),
            sheet_name,
            data: xl_sheet,
            search_points,
            range_start,
        })
    }
}
