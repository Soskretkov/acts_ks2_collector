use super::books::Book;
use super::tags::{TagAddressMap, TAG_INFO_ARRAY, TextCmp};
use crate::errors::Error;
use calamine::{DataType, Range, Reader};
use std::path::PathBuf;
pub struct Sheet {
    pub path: PathBuf,
    pub sheet_name: String,
    pub data: Range<DataType>,
    pub tag_address_map: TagAddressMap,
    pub range_start: (usize, usize),
}

impl<'a> Sheet {
    pub fn new(
        // разработчики Calamine делают зачем-то &mut self в функции worksheet_range(&mut self, name: &str),
        // из-за этого workbook приходится держать мутабельным, хотя этот код его менять вовсе не собирается
        // (из-за мутабельности workbook проблема при попытке множественных ссылок: можно только клонировать)
        workbook: &'a mut Book,
        user_entered_sh_name: &'a str,
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

        let mut tag_address_map = TagAddressMap::new();

        let mut limited_cell_iterator = xl_sheet.used_cells();
        let mut found_cell;

        for tag_info in TAG_INFO_ARRAY {
            let mut non_limited_cell_iterator = xl_sheet.used_cells();

            // Для обязательных тегов расходуемый итератор обеспечит валидацию очередности вохождения тегов
            // (например, "Стройку" мы ожидаем выше "Объекта, не наоборот).
            // Необязательным тегам расходуемый итератор не подходит, т.к. необязательный тег при отсутсвии израсходует итератор
            let iterator = if tag_info.is_required {
                &mut limited_cell_iterator
            } else {
                &mut non_limited_cell_iterator
            };

            found_cell = iterator.find(|cell| match cell.2.get_string() {
                Some(cell_content) => {
                    let cell_content = if tag_info.match_case {
                        cell_content.to_owned()
                    } else {
                        cell_content.to_lowercase()
                    };
                    let search_content = if tag_info.match_case {
                        tag_info.id.as_str().to_owned()
                    } else {
                        tag_info.id.as_str().to_lowercase()
                    };
        
                    match tag_info.look_at {
                        TextCmp::Whole => cell_content == search_content,
                        TextCmp::Part => cell_content.contains(&search_content),
                        TextCmp::StartsWith => cell_content.starts_with(&search_content),
                        TextCmp::EndsWith => cell_content.ends_with(&search_content),
                    }
                }
                None => false,
            });

            if let Some((row, col, _)) = found_cell {
                tag_address_map.insert(tag_info.id, (row, col));
            }
        }

        // Валидация на полноту данных: выше итератор расходующий ячейки и если хоть один поиск провалился, то это преждевременно
        // потребит все ячейки и извлечение по тегу последней строки в SEARCH_TAGS гарантированно провалится
        let validation_tag = TAG_INFO_ARRAY
            .iter()
            .filter(|search_tag| search_tag.is_required)
            .last()
            .ok_or_else(|| Error::InternalLogic {
                tech_descr: "Массив с тегами для поиска пуст".to_string(),
                err: None,
            })?
            .id;

        tag_address_map
            .get(&validation_tag)
            // нужно подменить штатную ошибку на ошибку валидации
            .or_else(|_| {
                Err(Error::SheetNotContainAllNecessaryData {
                    file_path: &workbook.path,
                })
            })?;

        let range_start = (sheet_start_coords.0 as usize, sheet_start_coords.1 as usize);

        Ok(Sheet {
            path: workbook.path.clone(),
            sheet_name,
            data: xl_sheet,
            tag_address_map,
            range_start,
        })
    }
}
