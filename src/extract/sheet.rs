use super::books::Book;
use super::tags::{Column, Row, TagAddressMap, TagID, TextCmp, TAG_INFO_ARRAY};
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
        // разработчики Calamine зачем-то требуют передать &mut в функцию worksheet_range(&mut self, name: &str),
        // из-за этого workbook приходится держать мутабельным, хотя этот код его менять не собирается
        // (это создает ряд проблем, в частности из-за мутабельности workbook приходится чаще клонировать)
        mut workbook: Book,
        user_entered_sh_name: &'a str,
    ) -> Result<Sheet, Error<'a>> {
        let entered_sh_name_lowercase = user_entered_sh_name.to_lowercase();

        let sheet_name = workbook
            .data
            .sheet_names()
            .iter()
            .find(|name| name.to_lowercase() == entered_sh_name_lowercase)
            .ok_or({
                let path_clone = workbook.path.clone();
                Error::CalamineSheetOfTheBookIsUndetectable {
                    file_path: path_clone,
                    sh_name_for_search: user_entered_sh_name,
                    sh_names: workbook.data.sheet_names().to_owned(),
                }
            })?
            .clone();

        let xl_sheet = workbook
            .data
            .worksheet_range(&sheet_name)
            .ok_or({
                let path_clone = workbook.path.clone();
                Error::CalamineSheetOfTheBookIsUndetectable {
                    file_path: path_clone,
                    sh_name_for_search: user_entered_sh_name,
                    sh_names: workbook.data.sheet_names().to_owned(),
                }
            })?
            .map_err(|error| {
                let path_clone = workbook.path.clone();
                Error::CalamineSheetOfTheBookIsUnreadable {
                    file_path: path_clone,
                    sh_name: sheet_name.to_owned(),
                    err: error,
                }
            })?;

        // при ошибки передается точное имя листа с учетом регистра (не используем ввод пользователя)
        let sheet_start_coords = xl_sheet.start().ok_or({
            let path_clone = workbook.path.clone();
            Error::EmptySheetRange {
                file_path: path_clone,
                sh_name: sheet_name.to_owned(),
            }
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
            .map_err(|_| {
                let path_clone = workbook.path.clone();
                Error::SheetNotContainAllNecessaryData {
                    file_path: path_clone,
                }
            })?;

        let range_start = (sheet_start_coords.0 as usize, sheet_start_coords.1 as usize);

        let result = Sheet {
            path: workbook.path.clone(),
            sheet_name,
            data: xl_sheet,
            tag_address_map,
            range_start,
        };

        check_row_type_alignment(&result)?;
        check_col_type_alignment(&result)?;

        Ok(result)
    }
}

fn check_row_type_alignment(sheet: &Sheet) -> Result<(), Error<'static>> {
    let mut valid_header_adr: Option<(usize, usize)> = None;
    let mut valid_header_tag_id: Option<TagID> = None;

    let filterd_tag_infos = TAG_INFO_ARRAY
        .into_iter()
        .filter(|tag_info| tag_info.group_by_row.is_some());

    for tag_info in filterd_tag_infos {
        let tag_adr = match sheet.tag_address_map.get(&tag_info.id) {
            Ok(adr) => adr,
            Err(err) => {
                if tag_info.is_required {
                    return Err(err);
                } else {
                    continue;
                }
            }
        };

        if let Some(group_name) = tag_info.group_by_row {
            // для устойчивости кода проверка сразу обоих на None
            if valid_header_adr.is_none() && valid_header_tag_id.is_none() {
                valid_header_adr = Some(*tag_adr);
                valid_header_tag_id = Some(tag_info.id);
                continue;
            }

            match group_name {
                Row::TableHeader => {
                    if valid_header_adr.map(|adr| adr.0) != Some(tag_adr.0) {
                        return Err(pack_into_error(
                            sheet,
                            valid_header_tag_id,
                            tag_info.id,
                            true,
                        ));
                    }
                }
            }
        }
    }

    Ok(())
}

fn check_col_type_alignment(sheet: &Sheet) -> Result<(), Error<'static>> {
    let mut valid_initial_adr: Option<(usize, usize)> = None;
    let mut valid_initial_tag_id: Option<TagID> = None;
    let mut valid_contract_adr: Option<(usize, usize)> = None;
    let mut valid_contract_tag_id: Option<TagID> = None;

    let filterd_tag_infos = TAG_INFO_ARRAY
        .into_iter()
        .filter(|tag_info| tag_info.group_by_col.is_some());

    for tag_info in filterd_tag_infos {
        let tag_adr = match sheet.tag_address_map.get(&tag_info.id) {
            Ok(adr) => adr,
            Err(err) => {
                if tag_info.is_required {
                    return Err(err);
                } else {
                    continue;
                }
            }
        };

        if let Some(group_name) = tag_info.group_by_col {
            match group_name {
                Column::Initial => {
                    if valid_initial_adr.is_none() && valid_initial_tag_id.is_none() {
                        valid_initial_adr = Some(*tag_adr);
                        valid_initial_tag_id = Some(tag_info.id);
                        continue;
                    }

                    if valid_initial_adr.map(|adr| adr.1) != Some(tag_adr.1) {
                        return Err(pack_into_error(
                            sheet,
                            valid_initial_tag_id,
                            tag_info.id,
                            true,
                        ));
                    }
                },
                Column::Contract => {
                    if valid_contract_adr.is_none() && valid_contract_tag_id.is_none() {
                        valid_contract_adr = Some(*tag_adr);
                        valid_contract_tag_id = Some(tag_info.id);
                        continue;
                    }

                    if valid_contract_adr.map(|adr| adr.1) != Some(tag_adr.1) {
                        return Err(pack_into_error(
                            sheet,
                            valid_contract_tag_id,
                            tag_info.id,
                            false,
                        ));
                    }
                }
            }
        }
    }

    Ok(())
}

fn pack_into_error(
    sheet: &Sheet,
    wrapped_first_tag: Option<TagID>,
    second_tag: TagID,
    is_row_alignment_check: bool,
) -> Error<'static> {
    let first_tag = match wrapped_first_tag {
        Some(tag) => tag,
        None => return Error::InternalLogic {
            tech_descr: "Отсутствует идентификатор тега для валидации смещения столбцов и строк в акте КС-2. Ожидалось, что идентификатор будет обязательно установлен в процессе валидации документа.".to_string(),
            err: None,
        },
    };

    let first_tag_adr_on_sheet = match sheet.tag_address_map.get(&first_tag) {
        Ok(adr) => (adr.0 + sheet.range_start.0 + 1, adr.1 + sheet.range_start.1 + 1),
        Err(err) => return err,
    };

    let second_tag_adr_on_sheet = match sheet.tag_address_map.get(&second_tag) {
        Ok(adr) => (adr.0 + sheet.range_start.0 + 1, adr.1 + sheet.range_start.1 + 1),
        Err(err) => return err,
    };

    Error::SheetMisalignment {
        is_row_alignment_check,
        first_tag_str: first_tag.as_str(),
        first_tag_adr_on_sheet,
        second_tag_str: second_tag.as_str(),
        second_tag_adr_on_sheet,
        file_path: sheet.path.clone(),
    }
}
