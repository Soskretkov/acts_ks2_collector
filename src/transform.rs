use crate::extract::*;
use calamine::{DataType, Reader};
use std::collections::HashMap;

impl Sheet {
    pub fn new<'a>(
        workbook: &'a mut Book,
        sheetname: &'static str,
        search_reference_points: &[(usize, Required, &'static str)],
        expected_sum_of_requir_col: usize,
    ) -> Result<Sheet, &'static str> {
        //) -> Result<Sheet, Box<dyn Error>> {
        let data = workbook.data.worksheet_range(sheetname).unwrap().unwrap();
        let mut search_points = HashMap::new();

        let mut temp_sh_iter = data.used_cells();
        for item in search_reference_points {
            let temp = temp_sh_iter
                .find(|x| x.2.get_string().unwrap_or("default") == item.2)
                .unwrap();
            search_points.insert(item.2, (temp.0, temp.1));
        }

        //Ниже значений на удаленность их столбцов чтобы гарантировать что найден нужный лист.
        let first_col = search_points.get("Стройка").unwrap().1;

        let (just_a_amount_requir_col, just_a_sum_requir_col) = search_reference_points
            .iter()
            .fold((0_usize, 0), |acc, item| match item.1 {
                Required::Y => (acc.0 + 1, acc.1 + search_points.get(item.2).unwrap().1),
                _ => acc,
            });

        match just_a_sum_requir_col - first_col * just_a_amount_requir_col
            == expected_sum_of_requir_col
        {
            true => {
                Ok(Sheet {
                    path: workbook.path.clone(),
                    sheetname,
                    data,
                    search_points,
                })
            }
            false => Err("Не найдена шапка КС-2"),
        }
    }
}

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

        let mut data_of_totals = Self::raw_totals(&sheet).unwrap(); //unwrap не требует обработки: функция возвращает только Ok вариант
        Self::renaming_totals(&mut data_of_totals);

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
        let stroika_adr = search_points.get("Стройка").unwrap(); //unwrap не требует обработки
        let object_adr = search_points.get("Объект").unwrap(); //unwrap не требует обработки
        let contrac_adr = search_points.get("Договор подряда").unwrap(); //unwrap не требует обработки
        let document_number_adr = search_points.get("Номер документа").unwrap(); //unwrap не требует обработки
        let workname_adr = search_points.get("Наименование работ и затрат").unwrap(); //unwrap не требует обработки

        let temp_vec: Vec<Option<(usize, usize)>> = NAMES_OF_HEADER.iter().fold(Vec::new(), |mut vec, shift| {

                let temp_cells_address: Option<(usize, usize)> = match shift {
                    (_, Some((point_name, (row, col)))) => {
                        let temp = match *point_name {
                            "Объект" => ((object_adr.0 as isize + *row as isize) as usize, (object_adr.1 as isize + *col as isize) as usize),
                            "Договор подряда" => ((contrac_adr.0 as isize + *row as isize) as usize, (contrac_adr.1 as isize + *col as isize) as usize),
                            "Номер документа" => ((document_number_adr.0 as isize + *row as isize) as usize, (document_number_adr.1 as isize + *col as isize) as usize),
                            "Наименование работ и затрат" => ((workname_adr.0 as isize + *row as isize) as usize, (workname_adr.1 as isize + *col as isize) as usize),
                            _ => unreachable!("Ошибка в логике программы, сообщающая о необходимости исправить код программы: ячейка в Excel с содержимым '{}' будет причиной подобных ошибок, пока не станет типом Required::Y чтобы обрабатываться", point_name),
                        };
                        Some(temp)
                    },
                    (target_name, _) => match *target_name {
                        "Исполнитель" => search_points.get("Исполнитель").map(|(row, col)| (*row, *col + 3)),
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
            .get("Стоимость материальных ресурсов (всего)")
            .unwrap(); //unwrap не требует обработки

        let total_row = sheet.data.get_size().0;
        let base_col = starting_col + 6;
        let current_col = starting_col + 9;

        let temp_vec_row = (starting_row..total_row).fold(Vec::new(), |mut acc, row| {
            let wrapped_row_name = &sheet.data[(row, starting_col)];
            if wrapped_row_name.is_string() {
                let base_price = &sheet.data[(row, base_col)];
                let current_price = &sheet.data[(row, current_col)];

                if base_price.is_float() || current_price.is_float() {
                    let row_name = wrapped_row_name.get_string().unwrap().to_string(); //unwrap не нужно обрабатывать: выше была проверка name.is_string
                                                                                       // if true {

                    let temp_total_row = TotalsRow {
                        name: row_name,
                        base_price: base_price.get_float(),
                        current_price: current_price.get_float(),
                    };
                    acc.push(temp_total_row);
                }
            }
            acc
        });

        Ok(temp_vec_row)
        // Err(format!("Ошибка повторяющихся строк в итогах акта: {} имеет строки с повторяющимися наименованиями затрат, таких строк {} шт.", sheet.path, len_diff))
    }

    fn renaming_totals(vec: &mut [TotalsRow]) {
        vec.iter_mut().fold(Vec::<&String>::new(), |mut uniq, row| {
            let mut new_name = row.name.clone();
            let mut counter = 1;
            while uniq.iter().any(|x| **x == new_name) {
                new_name = format!("{}, {}_{}", row.name, "duplicate", counter);
                counter += 1;
            }
            row.name = new_name;
            uniq.push(&row.name);
            uniq
        });
    }
}

#[derive(Debug, Clone)]
pub struct TotalsRow {
    pub name: String,
    pub base_price: Option<f64>,
    pub current_price: Option<f64>,
}

#[derive(Debug, Clone)]
pub enum DateVariant {
    String(String),
    Float(f64),
}
