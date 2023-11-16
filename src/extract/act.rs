use super::sheet::Sheet;
use super::tags::{TagAddressMap, TagArrayTools, TagID};
use crate::errors::Error;
use crate::types::XlDataType;
use calamine::DataType;

#[derive(Debug, Clone)]
pub struct CellCoords {
    pub row: (TagID, i8),
    pub col: (TagID, i8),
}

#[derive(Debug, Clone)]
pub struct DesiredCell {
    pub name: &'static str,
    pub cell_coords: Option<CellCoords>,
}

// Some смещение безопасно задавать только обязательным тегом
#[rustfmt::skip]
const DESIRED_CELLS_ARRAY: [DesiredCell; 16] = [
    DesiredCell{name:"Исполнитель",                  cell_coords: None},
    DesiredCell{name:"Глава",                        cell_coords: None},
    DesiredCell{name:"Глава наименование",           cell_coords: None},
    DesiredCell{name:"Объект",                       cell_coords: Some(CellCoords{row: (TagID::Объект, 0), col: (TagID::НаименованиеРаботИЗатрат, 0)})},
    DesiredCell{name:"Договор №",                    cell_coords: Some(CellCoords{row: (TagID::ДоговорПодряда, 0), col: (TagID::ДоговорПодряда, 2)})},
    DesiredCell{name:"Договор дата",                 cell_coords: Some(CellCoords{row: (TagID::ДоговорПодряда, 1), col: (TagID::ДоговорПодряда, 2)})},
    DesiredCell{name:"Смета №",                      cell_coords: Some(CellCoords{row: (TagID::ДоговорПодряда, 0), col: (TagID::Стройка, 0)})},
    DesiredCell{name:"Смета наименование",           cell_coords: Some(CellCoords{row: (TagID::ДоговорПодряда, 1), col: (TagID::Стройка, 0)})},
    DesiredCell{name:"По смете в ц.2000г.",          cell_coords: Some(CellCoords{row: (TagID::ДопСоглашение, 0), col: (TagID::НомерДокумента, 0)})},
    DesiredCell{name:"Выполнение работ в ц.2000г.",  cell_coords: Some(CellCoords{row: (TagID::ДопСоглашение, 1), col: (TagID::НомерДокумента, 0)})},
    DesiredCell{name:"Акт №",                        cell_coords: Some(CellCoords{row: (TagID::НомерДокумента, 2), col: (TagID::НомерДокумента, 0)})},
    DesiredCell{name:"Акт дата",                     cell_coords: Some(CellCoords{row: (TagID::НомерДокумента, 2), col: (TagID::НомерДокумента, 4)})},
    DesiredCell{name:"Отчетный период начало",       cell_coords: Some(CellCoords{row: (TagID::НомерДокумента, 2), col: (TagID::НомерДокумента, 5)})},
    DesiredCell{name:"Отчетный период окончание",    cell_coords: Some(CellCoords{row: (TagID::НомерДокумента, 2), col: (TagID::НомерДокумента, 6)})},
    DesiredCell{name:"Метод расчета",                cell_coords: Some(CellCoords{row: (TagID::НаименованиеРаботИЗатрат, -1), col: (TagID::Стройка, 0)})},
    DesiredCell{name:"Затраты труда, чел.-час",      cell_coords: None}, // небезопасно задать как Some, необязательные теги требуют особого подхода
];

#[derive(Debug, Clone)]
pub struct TotalsRow {
    pub name: String,
    pub base_price: Vec<Option<f64>>,
    pub curr_price: Vec<Option<f64>>,
    pub row_number: Vec<usize>,
}

#[derive(Debug, Clone)]
pub struct Act {
    pub path: String,
    pub sheetname: String,
    pub names_of_header: &'static [DesiredCell],
    pub data_of_header: Vec<Option<XlDataType>>,
    pub data_of_totals: Vec<TotalsRow>,
    pub start_row_of_totals: usize,
}

impl Act {
    pub fn new(sheet: Sheet) -> Result<Act, Error<'static>> {
        let header_addresses = Self::calculate_header_cell_addresses(&sheet.tag_address_map)?;
        // println!("{:#?}", header_addresses);
        let data_of_header: Vec<Option<XlDataType>> = header_addresses
            .iter()
            .map(|address| {
                // println!("Обрабатывается адрес: {:?}", address);

                match address {
                    Some(adr) => match &sheet.data[*adr] {
                        DataType::DateTime(x) => Some(XlDataType::Float(*x)),
                        DataType::Float(x) => Some(XlDataType::Float(*x)),
                        DataType::String(x) => {
                            Some(XlDataType::String(x.trim().replace("\r\n", "")))
                        }
                        _ => None,
                    },
                    None => None,
                }
            })
            .collect();

        let (start_row_of_totals_in_range, start_col_of_totals_in_range) =
            *sheet
                .tag_address_map
                .get(&TagID::СтоимостьМатериальныхРесурсовВсего)?;

        let data_of_totals = Self::get_totals(
            &sheet,
            (start_row_of_totals_in_range, start_col_of_totals_in_range),
        )?;

        let start_row_of_totals = start_row_of_totals_in_range + sheet.range_start.0 + 1;
        Ok(Act {
            path: sheet.path.to_string_lossy().to_string(),
            sheetname: sheet.sheet_name,
            names_of_header: &DESIRED_CELLS_ARRAY,
            data_of_header,
            data_of_totals,
            start_row_of_totals,
        })
    }
    fn calculate_header_cell_addresses(
        tag_address_map: &TagAddressMap,
    ) -> Result<Vec<Option<(usize, usize)>>, Error<'static>> {
        let stroika_adr = tag_address_map.get(&TagID::Стройка)?;
        let object_adr = tag_address_map.get(&TagID::Объект)?;
        let naimenov_rabot_i_zatr_adr = tag_address_map.get(&TagID::НаименованиеРаботИЗатрат)?;
        // Предполагается что строка с Главой должна четко находиться между "Стройкой" и "Объектом"
        let is_valid_glava = stroika_adr.0 + 2 == object_adr.0;
        let mut vec: Vec<Option<(usize, usize)>> = Vec::new();

        for item in DESIRED_CELLS_ARRAY {
            let temp_cells_address: Option<(usize, usize)> = match item {
                DesiredCell {
                    name: _,
                    cell_coords: Some(cell_coords_struct),
                } => Some(calculate_cell_adr_by_coords(tag_address_map, cell_coords_struct)?),
                // рукав обработывает ячейки, поиск которых основан на необязательных тегах (индивидуальная логика)
                DesiredCell { name, .. } => match name {
                    "Исполнитель" => tag_address_map
                        .get(&TagID::Исполнитель)
                        .ok()
                        .map(|(row_adr, _)| (*row_adr, naimenov_rabot_i_zatr_adr.1)),
                    "Глава" => {
                        if is_valid_glava {
                            Some((stroika_adr.0 + 1, stroika_adr.1))
                        } else {
                            None
                        }
                    }
                    "Глава наименование" => {
                        if is_valid_glava {
                            Some((stroika_adr.0 + 1, naimenov_rabot_i_zatr_adr.1))
                        } else {
                            None
                        }
                    }
                    "Затраты труда, чел.-час" => {
                        let ttl = tag_address_map.get(&TagID::ИтогоПоАкту).ok();
                        let ztr = tag_address_map.get(&TagID::ЗтрВсего).ok();

                        match (ttl, ztr) {
                            (Some(ttl_adr), Some(ztr_adr)) => Some((ttl_adr.0, ztr_adr.1)),
                            _ => None,
                        }
                    }
                    _ => {
                        return Err(Error::InternalLogic {
                            tech_descr: format!(
                                r#"Match не имеет рукав, обрабатывающий элемент акта "{}""#,
                                name
                            ),
                            err: None,
                        })
                    }
                },
            };

            vec.push(temp_cells_address);
        }
        Ok(vec)
    }

    fn get_totals(
        sheet: &Sheet,
        totals_start_adr: (usize, usize),
    ) -> Result<Vec<TotalsRow>, Error<'static>> {
        let (totals_start_row, totals_start_col) = totals_start_adr;
        let total_row = sheet.data.get_size().0;
        let base_col = sheet.tag_address_map.get(&TagID::СтоимостьВЦенах2001)?.1;
        let current_col = sheet.tag_address_map.get(&TagID::СтоимостьВТекущихЦенах)?.1;

        let mut blank_row_flag = false;
        let mut totals_row_vec = Vec::<TotalsRow>::new();

        for row in totals_start_row..total_row {
            let row_data_type = &sheet.data[(row, totals_start_col)];
            if row_data_type.is_string() {
                let base_price = &sheet.data[(row, base_col)];
                let current_price = &sheet.data[(row, current_col)];

                //Если пустых ячеек вместо имени еще не встречалось, то собираем данные независимо от наличия цены.
                //Ситуация меняется если встретилось первое пустое имя: теперь потребуется и имя и цена (перестраховка на случай случайных пустых строк)
                if !blank_row_flag || base_price.is_float() || current_price.is_float() {
                    let row_name = row_data_type
                        .get_string()
                        .ok_or_else(|| Error::InternalLogic {
                            tech_descr: "При работе с ячейкой в итогах акта ожидался валидированный строковый тип данных Excel.".to_string(),
                            err: None,
                        })?
                        .trim()
                        .replace("\r\n", "");

                    match totals_row_vec
                        .iter_mut()
                        .find(|object| object.name == row_name)
                    {
                        Some(x) => {
                            x.base_price.push(base_price.get_float());
                            x.curr_price.push(current_price.get_float());
                            x.row_number.push(sheet.range_start.0 + row + 1);
                        }
                        None => {
                            let temp_total_row = TotalsRow {
                                name: row_name,
                                base_price: vec![base_price.get_float()],
                                curr_price: vec![current_price.get_float()],
                                row_number: vec![sheet.range_start.0 + row + 1],
                            };
                            totals_row_vec.push(temp_total_row);
                        }
                    }
                }
            } else if !blank_row_flag {
                blank_row_flag = true;
            }
        }

        Ok(totals_row_vec)
    }
}

fn calculate_cell_adr_by_coords(
    tag_address_map: &TagAddressMap,
    cell_coords: CellCoords,
) -> Result<(usize, usize), Error<'static>> {
    let row_tag_adr = tag_address_map.get(&cell_coords.row.0)?;
    let col_tag_adr = tag_address_map.get(&cell_coords.col.0)?;
    let row_ofst = cell_coords.row.1;
    let col_ofst = cell_coords.col.1;

    let row_result = row_tag_adr
        .0
        .try_into()
        .map_err(|err| Error::NumericConversion {
            tech_descr: format!(
                r#"Не удалась конвертация типа usize с значением "{}" в тип isize."#,
                row_tag_adr.0
            ),
            err: Box::new(err),
        })
        .and_then(|row: isize| {
            row.checked_add(row_ofst as isize)
                .ok_or_else(|| Error::NumericOverflowError {
                    tech_descr: format!(
                        "Переполнение при сложении типа isize ({}) с типом isize ({}).",
                        row, row_ofst
                    ),
                })
        })
        .map(|row| row as usize);

    let col_result = col_tag_adr
        .1
        .try_into()
        .map_err(|err| Error::NumericConversion {
            tech_descr: format!(
                r#"Не удалась конвертация типа usize с значением "{}" в тип isize."#,
                col_tag_adr.1
            ),
            err: Box::new(err),
        })
        .and_then(|col: isize| {
            col.checked_add(col_ofst as isize)
                .ok_or_else(|| Error::NumericOverflowError {
                    tech_descr: format!(
                        "Переполнение при сложении типа isize ({}) с типом isize ({}).",
                        col, col_ofst
                    ),
                })
        })
        .map(|c| c as usize);

    match (row_result, col_result) {
        (Ok(row), Ok(col)) => Ok((row, col)),
        (Err(e), _) | (_, Err(e)) => return Err(e),
    }
}
