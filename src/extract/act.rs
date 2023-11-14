use super::sheet::Sheet;
use super::tags::{TagAddressMap, TagArrayTools, TagID};
use crate::constants::{BASE_PRICE_COLUMN_OFFSET, CURRENT_PRICE_COLUMN_OFFSET};
use crate::errors::Error;
use crate::types::XlDataType;
use calamine::DataType;

#[derive(Debug, Clone)]
pub struct DesiredCell {
    pub name: &'static str,
    pub offset: Option<(TagID, (i8, i8))>,
}

#[rustfmt::skip]
const DESIRED_CELLS_ARRAY: [DesiredCell; 16] = [
    DesiredCell{name:"Исполнитель",                  offset: Some((TagID::Исполнитель,              (0, 3)))},
    DesiredCell{name:"Глава",                        offset: None},
    DesiredCell{name:"Глава наименование",           offset: None},
    DesiredCell{name:"Объект",                       offset: Some((TagID::Объект,                   (0, 3)))},
    DesiredCell{name:"Договор №",                    offset: Some((TagID::ДоговорПодряда,           (0, 2)))},
    DesiredCell{name:"Договор дата",                 offset: Some((TagID::ДоговорПодряда,           (1, 2)))},
    DesiredCell{name:"Смета №",                      offset: Some((TagID::ДоговорПодряда,           (0, -9)))},
    DesiredCell{name:"Смета наименование",           offset: Some((TagID::ДоговорПодряда,           (1, -9)))},
    DesiredCell{name:"По смете в ц.2000г.",          offset: Some((TagID::ДопСоглашение,            (0, -4)))},
    DesiredCell{name:"Выполнение работ в ц.2000г.",  offset: Some((TagID::ДопСоглашение,            (1, -4)))},
    DesiredCell{name:"Акт №",                        offset: Some((TagID::НомерДокумента,           (2, 0)))},
    DesiredCell{name:"Акт дата",                     offset: Some((TagID::НомерДокумента,           (2, 4)))},
    DesiredCell{name:"Отчетный период начало",       offset: Some((TagID::НомерДокумента,           (2, 5)))},
    DesiredCell{name:"Отчетный период окончание",    offset: Some((TagID::НомерДокумента,           (2, 6)))},
    DesiredCell{name:"Метод расчета",                offset: Some((TagID::НаименованиеРаботИЗатрат, (-1, -3)))},
    DesiredCell{name:"Затраты труда, чел.-час",      offset: None},
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
    pub names_of_header: &'static [DesiredCell; 16],
    pub data_of_header: Vec<Option<XlDataType>>,
    pub data_of_totals: Vec<TotalsRow>,
    pub start_row_of_totals: usize,
}

impl Act {
    pub fn new(sheet: Sheet) -> Result<Act, Error<'static>> {
        let header_addresses = Self::calculate_header_cell_addresses(&sheet.tag_address_map)?;
        let data_of_header: Vec<Option<XlDataType>> = header_addresses
            .iter()
            .map(|address| match address {
                Some(x) => match &sheet.data[*x] {
                    DataType::DateTime(x) => Some(XlDataType::Float(*x)),
                    DataType::Float(x) => Some(XlDataType::Float(*x)),
                    DataType::String(x) => Some(XlDataType::String(x.trim().replace("\r\n", ""))),
                    _ => None,
                },
                None => None,
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
        let is_valid_glava = stroika_adr.0 + 2 == object_adr.0;
        let mut vec: Vec<Option<(usize, usize)>> = Vec::new();

        for item in DESIRED_CELLS_ARRAY {
            let temp_cells_address: Option<(usize, usize)> = match item {
                DesiredCell {
                    name: _,
                    offset: Some((tag_id, (row, col))),
                } => {
                    let tag_info = TagArrayTools::get_tag_info_by_id(tag_id)?;
                    let wrapped_adr = tag_address_map.get(&tag_id);

                    // Возвращаем ошибку, если тег обязателен, но адрес не найден
                    if wrapped_adr.is_err() && tag_info.is_required {
                        return Err(Error::InternalLogic {
                            tech_descr: format!(
                                r#"Хешкарта не содержит ключ "{}""#,
                                tag_id.as_str()
                            ),
                            err: None,
                        });
                    }

                    wrapped_adr.ok().map(|adr| {
                        (
                            (adr.0 as isize + row as isize) as usize,
                            (adr.1 as isize + col as isize) as usize,
                        )
                    })
                }
                DesiredCell { name, .. } => match name {
                    "Глава" => {
                        if is_valid_glava {
                            Some((stroika_adr.0 + 1, stroika_adr.1))
                        } else {
                            None
                        }
                    }
                    "Глава наименование" => {
                        if is_valid_glava {
                            Some((stroika_adr.0 + 1, stroika_adr.1 + 3))
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
        first_row_address: (usize, usize),
    ) -> Result<Vec<TotalsRow>, Error<'static>> {
        let (starting_row, starting_col) = first_row_address;
        let total_row = sheet.data.get_size().0;
        let base_col = starting_col + BASE_PRICE_COLUMN_OFFSET;
        let current_col = starting_col + CURRENT_PRICE_COLUMN_OFFSET;

        let mut blank_row_flag = false;
        let mut totals_row_vec = Vec::<TotalsRow>::new();

        for row in starting_row..total_row {
            let row_data_type = &sheet.data[(row, starting_col)];
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
