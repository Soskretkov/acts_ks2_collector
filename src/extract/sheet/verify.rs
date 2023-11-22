use super::Sheet;
use crate::errors::Error;
use crate::extract::tags::{Column, Row, TagID, TAG_INFO_ARRAY};
use crate::shared::utils;

pub fn check_row_type_alignment(sheet: &Sheet) -> Result<(), Error<'static>> {
    let is_row_algmnt_check = true;
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
                            is_row_algmnt_check,
                        ));
                    }
                }
            }
        }
    }

    Ok(())
}

pub fn check_col_type_alignment(sheet: &Sheet) -> Result<(), Error<'static>> {
    let is_row_algmnt_check = false;
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
                            is_row_algmnt_check,
                        ));
                    }
                }
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
                            is_row_algmnt_check,
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
    is_row: bool,
) -> Error<'static> {
    let first_tag = match wrapped_first_tag {
        Some(tag) => tag,
        None => return Error::InternalLogic {
            tech_descr: "Отсутствует идентификатор тега для валидации смещения столбцов и строк в акте КС-2. Ожидалось, что идентификатор будет обязательно установлен в процессе валидации документа.".to_string(),
            err: None,
        },
    };

    let fst_tag_index_on_sheet = match get_xl_column_letter_or_row_idx(sheet, first_tag, is_row) {
        Ok(idx) => idx,
        Err(err) => return err,
    };

    let snd_tag_index_on_sheet = match get_xl_column_letter_or_row_idx(sheet, second_tag, is_row) {
        Ok(idx) => idx,
        Err(err) => return err,
    };

    Error::SheetMisalignment {
        is_row_algmnt_check: is_row,
        fst_tag_str: first_tag.as_str(),
        fst_tag_index_on_sheet,
        snd_tag_str: second_tag.as_str(),
        snd_tag_index_on_sheet,
        file_path: sheet.path.clone(),
    }
}

fn get_xl_column_letter_or_row_idx(
    sheet: &Sheet,
    tag: TagID,
    is_row: bool,
) -> Result<String, Error<'static>> {
    let usize_zero_based_idx_on_sheet = match sheet.tag_address_map.get(&tag) {
        Ok(adr) => match is_row {
            true => adr.0 + sheet.range_start.0,
            false => adr.1 + sheet.range_start.1,
        },
        Err(err) => return Err(err),
    };

    let tag_index_on_sheet = if is_row {
        // важно перейти в One-Based Indexing
        (usize_zero_based_idx_on_sheet + 1).to_string()
    } else {
        let u16_zero_based_idx_on_sheet: u16 =
            usize_zero_based_idx_on_sheet
                .try_into()
                .map_err(|err| Error::NumericConversion {
                    tech_descr: format!(
                        r#"Не удалась конвертация типа usize с значением "{}" в тип u16."#,
                        usize_zero_based_idx_on_sheet
                    ),
                    err: Box::new(err),
                })?;
        // у функции аргумент в Zero-Based Indexing
        utils::get_xl_column_letter(u16_zero_based_idx_on_sheet)
    };

    Ok(tag_index_on_sheet)
}
