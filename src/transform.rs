use crate::extract::*;
use calamine::DataType;
use std::collections::HashMap;

#[derive(Debug, Clone)]
pub struct Act {
    pub path: String,
    pub sheetname: &'static str,
    pub names_of_header: &'static [(&'static str, Option<(&'static str, (i8, i8))>); 15],
    pub data_of_header: Vec<Option<DateVariant>>,
    pub data_of_totals: Vec<TotalsRow>,
}

impl<'a> Act {
    pub fn new(sheet: Sheet) -> Result<Act, String> {
        let header_addresses = Self::cells_addreses_in_header(&sheet.search_points);
        let data_of_header: Vec<Option<DateVariant>> = header_addresses
            .iter()
            .map(|address| match address {
                Some(x) => match &sheet.data[*x] {
                    DataType::DateTime(x) => Some(DateVariant::Float(*x)),
                    DataType::Float(x) => Some(DateVariant::Float(*x)),
                    DataType::String(x) => Some(DateVariant::String(x.to_owned())),
                    _ => None,
                },
                None => None,
            })
            .collect();

        let data_of_totals = Self::raw_totals(&sheet).unwrap(); //unwrap не требует обработки: функция возвращает только Ok вариант

        Ok(Act {
            path: sheet.path,
            sheetname: sheet.sheetname,
            names_of_header: &NAMES_OF_HEADER,
            data_of_header,
            data_of_totals,
        })
    }
    fn cells_addreses_in_header(
        search_points: &HashMap<&'static str, (usize, usize)>,
    ) -> Vec<Option<(usize, usize)>> {
        let stroika_adr = search_points.get("стройка").unwrap(); //unwrap не требует обработки
        let object_adr = search_points.get("объект").unwrap(); //unwrap не требует обработки
        let contrac_adr = search_points.get("договор подряда").unwrap(); //unwrap не требует обработки
        let dopsogl_adr = search_points.get("доп. соглашение").unwrap(); //unwrap не требует обработки
        let document_number_adr = search_points.get("номер документа").unwrap(); //unwrap не требует обработки
        let workname_adr = search_points.get("наименование работ и затрат").unwrap(); //unwrap не требует обработки

        let temp_vec: Vec<Option<(usize, usize)>> = NAMES_OF_HEADER.iter().fold(Vec::new(), |mut vec, shift| {

                let temp_cells_address: Option<(usize, usize)> = match shift {
                    (_, Some((point_name, (row, col)))) => {
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
                    (target_name, _) => match *target_name {
                        "Исполнитель" => search_points.get("исполнитель").map(|(row, col)| (*row, *col + 3)),
                        "Глава" => match stroika_adr.0 + 2 == object_adr.0 {
                            true => Some((stroika_adr.0 + 1, stroika_adr.1)),
                            false => None,
                        }//Адрес возвращается только если между "Стройкой" и "Объектом" одна строка
                        "Глава наименование" => match stroika_adr.0 + 2 == object_adr.0 {
                            true => Some((stroika_adr.0 + 1, stroika_adr.1 + 3)),
                            false => None,
                        }//Адрес возвращается только если между "Стройкой" и "Объектом" одна строка
                        _ => None,
                    },
                };

                vec.push(temp_cells_address);
                vec

            });
        temp_vec
    }

    fn raw_totals(sheet: &Sheet) -> Result<Vec<TotalsRow>, String> {
        let (starting_row, starting_col) = *sheet
            .search_points
            .get("стоимость материальных ресурсов (всего)")
            .unwrap(); //unwrap не требует обработки

        let total_row = sheet.data.get_size().0;
        let base_col = starting_col + 6;
        let current_col = starting_col + 9;

        let temp_vec_row =
            (starting_row..total_row).fold(Vec::<TotalsRow>::new(), |mut acc, row| {
                let wrapped_row_name = &sheet.data[(row, starting_col)];
                if wrapped_row_name.is_string() {
                    let base_price = &sheet.data[(row, base_col)];
                    let current_price = &sheet.data[(row, current_col)];

                    if base_price.is_float() || current_price.is_float() {
                        let row_name = wrapped_row_name.get_string().unwrap().to_string(); //unwrap не нужно обрабатывать: выше была проверка name.is_string

                        match acc.iter_mut().find(|object| object.name == row_name) {
                            Some(x) => {
                                x.base_price.push(base_price.get_float());
                                x.current_price.push(current_price.get_float());
                                x.row_number.push(row);
                            }
                            None => {
                                let temp_total_row = TotalsRow {
                                    name: row_name,
                                    base_price: vec![base_price.get_float()],
                                    current_price: vec![current_price.get_float()],
                                    row_number: vec![row],
                                };
                                acc.push(temp_total_row);
                            }
                        }
                    }
                }
                acc
            });

        Ok(temp_vec_row)
        // Err(format!("Ошибка повторяющихся строк в итогах акта: {} имеет строки с повторяющимися наименованиями затрат, таких строк {} шт.", sheet.path, len_diff))
    }
}

#[derive(Debug, Clone)]
pub struct TotalsRow {
    pub name: String,
    pub base_price: Vec<Option<f64>>,
    pub current_price: Vec<Option<f64>>,
    pub row_number: Vec<usize>,
}

#[derive(Debug, Clone)]
pub enum DateVariant {
    String(String),
    Float(f64),
}
