use calamine::DataType;
use std::collections::HashMap;
use crate::errors::Error;
use crate::extract::{DesiredData, Sheet, DESIRED_DATA_ARRAY};
use crate::constants::{BASE_PRICE_COLUMN_OFFSET, CURRENT_PRICE_COLUMN_OFFSET};

#[derive(Debug, Clone)]
pub enum DataVariant {
    String(String),
    Float(f64),
}

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
    pub names_of_header: &'static [DesiredData; 16],
    pub data_of_header: Vec<Option<DataVariant>>,
    pub data_of_totals: Vec<TotalsRow>,
    pub start_row_of_totals: usize,
}

impl Act {
    pub fn new(sheet: Sheet) -> Result<Act, Error<'static>> {
        let header_addresses = Self::cells_addreses_in_header(&sheet.search_points);
        let data_of_header: Vec<Option<DataVariant>> = header_addresses
            .iter()
            .map(|address| match address {
                Some(x) => match &sheet.data[*x] {
                    DataType::DateTime(x) => Some(DataVariant::Float(*x)),
                    DataType::Float(x) => Some(DataVariant::Float(*x)),
                    DataType::String(x) => Some(DataVariant::String(x.trim().replace("\r\n", ""))),
                    _ => None,
                },
                None => None,
            })
            .collect();

        let hmap_key = "стоимость материальных ресурсов (всего)";
        let (start_row_of_totals_in_range, start_col_of_totals_in_range) = *sheet
            .search_points
            .get(hmap_key)
            .ok_or_else(|| Error::InternalLogic {
                tech_descr: format!("HashMap не содержит ключ: {}", hmap_key),
                err: None,
            })?;

        let data_of_totals = Self::get_totals(
            &sheet,
            (start_row_of_totals_in_range, start_col_of_totals_in_range),
        )?;

        let start_row_of_totals = start_row_of_totals_in_range + sheet.range_start.0 + 1;
        Ok(Act {
            path: sheet.path.to_string_lossy().to_string(),
            sheetname: sheet.sheet_name,
            names_of_header: &DESIRED_DATA_ARRAY,
            data_of_header,
            data_of_totals,
            start_row_of_totals,
        })
    }
    fn cells_addreses_in_header(
        search_points: &HashMap<&'static str, (usize, usize)>,
    ) -> Vec<Option<(usize, usize)>> {
        //unwrap не требует обработки (валидировано)
        let stroika_adr = search_points.get("стройка").unwrap();  
        let object_adr = search_points.get("объект").unwrap();
        let contrac_adr = search_points.get("договор подряда").unwrap();
        let dopsogl_adr = search_points.get("доп. соглашение").unwrap();
        let document_number_adr = search_points.get("номер документа").unwrap();
        let workname_adr = search_points.get("наименование работ и затрат").unwrap();

        let temp_vec: Vec<Option<(usize, usize)>> = DESIRED_DATA_ARRAY.iter().fold(Vec::new(), |mut vec, shift| {

                let temp_cells_address: Option<(usize, usize)> = match shift {
                    DesiredData{name: _, offset: Some((point_name, (row, col)))} => {
                        let temp = match *point_name {
                            "объект" => ((object_adr.0 as isize + *row as isize) as usize, (object_adr.1 as isize + *col as isize) as usize),
                            "договор подряда" => ((contrac_adr.0 as isize + *row as isize) as usize, (contrac_adr.1 as isize + *col as isize) as usize),
                            "доп. соглашение" => ((dopsogl_adr.0 as isize + *row as isize) as usize, (dopsogl_adr.1 as isize + *col as isize) as usize),
                            "номер документа" => ((document_number_adr.0 as isize + *row as isize) as usize, (document_number_adr.1 as isize + *col as isize) as usize),
                            "наименование работ и затрат" => ((workname_adr.0 as isize + *row as isize) as usize, (workname_adr.1 as isize + *col as isize) as usize),
                            _ => unreachable!("Ошибка в логике программы, сообщающая о необходимости исправления программного кода: ячейка в Excel с содержимым '{}' будет причиной подобных ошибок, пока не станет типом Required::Y, подлежащим обработке", point_name),
                        };
                        Some(temp)
                    },
                    DesiredData{name: content, ..} => match *content {
                        "Исполнитель" => search_points.get("исполнитель").map(|(row, col)| (*row, *col + 3)),
                        "Глава" => match stroika_adr.0 + 2 == object_adr.0 {
                            true => Some((stroika_adr.0 + 1, stroika_adr.1)),
                            false => None,
                        }//Адрес возвращается только если между "Стройкой" и "Объектом" одна строка
                        "Глава наименование" => match stroika_adr.0 + 2 == object_adr.0 {
                            true => Some((stroika_adr.0 + 1, stroika_adr.1 + 3)),
                            false => None,
                        }
                        "Затраты труда, чел.-час" => {
                            let ttl = search_points.get("итого по акту:");
                            let ztr = search_points.get("зтр всего чел.-час");

                            if ztr.is_some() & ttl.is_some() {
                                Some((ttl.unwrap().0, ztr.unwrap().1))
                            } else {
                                None
                            }
                        }
                        _ => None,
                    },
                };

                vec.push(temp_cells_address);
                vec

            });
        temp_vec
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
