use crate::transform::{Act, DataVariant, TotalsRow};
use ks2_etl::variant_eq;
use regex::Regex;
use xlsxwriter::{DateTime, Format, FormatAlignment, Workbook, Worksheet, RowColOptions};

#[derive(Debug)]
pub struct OutputData {
    pub rename: Option<&'static str>,
    pub moving: Moving,
    pub expected_columns: u16,
    pub source: Source,
}
#[derive(Debug, Clone, PartialEq)]

pub enum Moving {
    No,
    Yes,
    Del,
}

#[derive(Debug, PartialEq, Clone)]
pub enum Matches {
    Exact,
    Contains,
}

// Четыре вида данных на выходе: в готовом виде в шапке, в готов виде в итогах акта (2 варанта), и нет готовых (нужно расчитать программой):
#[derive(Debug, PartialEq)]
pub enum Source {
    InTableHeader(&'static str),
    AtCurrPrices(String, Matches),
    AtBasePrices(String, Matches),
    Calculate(&'static str),
}

#[derive(Debug)]
pub struct PrintPart {
    vector: Vec<OutputData>,
    number_of_columns: u16,
}
impl PrintPart {
    pub fn new(vector: Vec<OutputData>) -> Option<PrintPart> {
        let number_of_columns = Self::count_col(&vector);

        Some(PrintPart {
            vector,
            number_of_columns,
        })
    }
    pub fn get_number_of_columns(&self) -> u16 {
        self.number_of_columns
    }

    pub fn get_index_and_address_by_columns(
        &self,
        kind: &str,
        name: &str,
        matches: Matches,
    ) -> Option<(usize, u16)> {
        let src = match kind {
            "base" => Source::AtBasePrices("".to_string(), matches.clone()),
            "curr" => Source::AtCurrPrices("".to_string(), matches.clone()),
            "calc" => Source::Calculate(""),
            "header" => Source::InTableHeader(""),
            _ => panic!(),
        };

        let mut counter = 0;
        let mut index = 0;
        for outputdata in self.vector.iter() {
            match outputdata {
                OutputData {
                    source: Source::Calculate(text) | Source::InTableHeader(text),
                    ..
                } if ks2_etl::variant_eq(&outputdata.source, &src) && &name == text => {
                    return Some((index, counter));
                }
                OutputData {
                    source: Source::AtBasePrices(text, m) | Source::AtCurrPrices(text, m),
                    ..
                } if ks2_etl::variant_eq(&outputdata.source, &src)
                    && variant_eq(m, &matches)
                    && m == &Matches::Exact
                    && name == text =>
                {
                    return Some((index, counter));
                }
                OutputData {
                    source: Source::AtBasePrices(text, m) | Source::AtCurrPrices(text, m),
                    ..
                } if ks2_etl::variant_eq(&outputdata.source, &src)
                    && variant_eq(m, &matches)
                    && m == &Matches::Contains
                    && name.contains(text) =>
                {
                    return Some((index, counter));
                }
                OutputData { moving: mov, .. } => {
                    if mov == &Moving::No || mov == &Moving::Yes {
                        index += 1;
                        counter += outputdata.expected_columns;
                    }
                }
            };
        }
        None
    }

    fn count_col(vector: &[OutputData]) -> u16 {
        vector
            .iter()
            .fold(0, |acc, outputdata| match outputdata.moving {
                Moving::No => acc + outputdata.expected_columns,
                Moving::Yes => acc + outputdata.expected_columns,
                _ => acc,
            })
    }
}
#[test]
fn PrintPart_test() {
    #[rustfmt::skip]
        let vec_to_test = vec![
            OutputData{rename: None,                           moving: Moving::No,  expected_columns: 1,  source: Source::InTableHeader("Объект")},
            OutputData{rename: None,                           moving: Moving::Yes, expected_columns: 2,  source: Source::AtBasePrices("Накладные расходы".to_string(), Matches::Exact)},
            OutputData{rename: None,                           moving: Moving::Yes, expected_columns: 3,  source: Source::AtBasePrices("Эксплуатация машин".to_string(), Matches::Exact)},
            OutputData{rename: None,                           moving: Moving::Yes, expected_columns: 4,  source: Source::AtCurrPrices("Накладные расходы".to_string(), Matches::Exact)},
            OutputData{rename: None,                           moving: Moving::Yes, expected_columns: 5,  source: Source::AtCurrPrices("Накладные".to_string(), Matches::Contains)},
            OutputData{rename: Some("РЕНЕЙМ................"), moving: Moving::No,  expected_columns: 6,  source: Source::AtCurrPrices("Производство работ в зимнее время 4%".to_string(), Matches::Exact)},
            OutputData{rename: Some("УДАЛИТЬ..............."), moving: Moving::Del, expected_columns: 99, source: Source::AtBasePrices("Производство работ в зимнее время 4%".to_string(), Matches::Exact)},
            OutputData{rename: None,                           moving: Moving::Yes, expected_columns: 8,  source: Source::AtCurrPrices("Стоимость материальных ресурсов (всего)".to_string(), Matches::Exact)},
        ];
    let printpart = PrintPart::new(vec_to_test).unwrap();

    assert_eq!(&29, &printpart.get_number_of_columns());
    assert_eq!(
        Some((6, 21)),
        printpart.get_index_and_address_by_columns(
            "curr",
            "Стоимость материальных ресурсов (всего)",
            Matches::Exact
        )
    );
    assert_eq!(
        Some((4, 10)),
        printpart.get_index_and_address_by_columns("curr", "Накладные расходы", Matches::Contains)
    );
}
pub struct Report {
    pub book: xlsxwriter::Workbook,
    pub part_main: PrintPart,
    pub part_base: Option<PrintPart>,
    pub part_curr: Option<PrintPart>,
    pub skip_row: u32,
    pub body_size: u32,
}

impl<'a> Report {
    pub fn new(wb: xlsxwriter::Workbook) -> Report {
        // Нужно чтобы код назначал длину таблицы по горизонтали в зависимости от количества строк в итогах (обычно итоги имеют 17 строк,
        // но если какой-то акт имеет 16, 18, 0 или, скажем, 40 строк в итогах, то нужна какая-то логика, чтобы соотнести эти 40 строк одного акта
        // с 17 строками других актов. Нужно решение, как не сокращать эти 40 строк до 17 стандартных и выдать информацию пользователю без потерь.
        // Данные делятся на ожидаемые (им порядок можно сразу задать) и случайные.
        // Ниже массив, содержащий информацию о колонках, которые мы ожидаем получить из актов, здесь будем задавать порядок.
        // Позиция в массиве будет соответсвовать столбцу выходной формы (это крайние левые столбцы шапки):

        #[rustfmt::skip]
        let main_list: Vec<OutputData> = vec![
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1,  source: Source::Calculate("Папка (ссылка)")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1,  source: Source::Calculate("Файл (ссылка)")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1,  source: Source::Calculate("Акт №")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1,  source: Source::InTableHeader("Акт дата")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1,  source: Source::InTableHeader("Исполнитель")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1,  source: Source::Calculate("Глава")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1,  source: Source::InTableHeader("Объект")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1,  source: Source::InTableHeader("Договор №")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1,  source: Source::InTableHeader("Договор дата")},
            // OutputData{rename: None,                                            moving: Moving::Yes,   expected_columns: 1, source: Source::AtBasePrices("Стоимость материальных ресурсов (всего)", Matches::Exact)},
            // OutputData{rename: Some("Восстание машин"),                         moving: Moving::No, expected_columns: 1, source: Source::AtBasePrices("Эксплуатация машин", Matches::Exact)},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1,  source: Source::Calculate("Смета №")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1,  source: Source::InTableHeader("Смета наименование")},
            OutputData{rename: Some("По смете в ц.2000г., руб."),           moving: Moving::No, expected_columns: 1,  source: Source::Calculate("По смете в ц.2000г.")},
            OutputData{rename: Some("Выполнение работ в ц.2000г., руб."),   moving: Moving::No, expected_columns: 1,  source: Source::Calculate("Выполнение работ в ц.2000г.")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1,  source: Source::InTableHeader("Отчетный период начало")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1,  source: Source::InTableHeader("Отчетный период окончание")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1,  source: Source::InTableHeader("Метод расчета")},
            OutputData{rename: None,                                        moving: Moving::Del, expected_columns: 1, source: Source::AtBasePrices("Всего с НР и СП (тек".to_string(), Matches::Contains)},
            OutputData{rename: None,                                        moving: Moving::Del, expected_columns: 1, source: Source::AtCurrPrices("Всего с НР и СП (баз".to_string(), Matches::Contains)},
            OutputData{rename: None,                                        moving: Moving::Del, expected_columns: 1, source: Source::AtBasePrices("Итого с К = 1".to_string(), Matches::Exact)},
            OutputData{rename: None,                                        moving: Moving::Del, expected_columns: 1, source: Source::AtCurrPrices("Итого с К = 1".to_string(), Matches::Exact)},
            // OutputData{rename: Some("РЕНЕЙМ................"),                  moving: Moving::No, expected_columns: 1, source: Source::AtBasePrices("Производство работ в зимнее время 4%", Matches::Exact)},
            // OutputData{rename: None,                                            moving: Moving::Yes, expected_columns: 1, source: Source::AtBasePrices("ы", Matches::Contains)},
        ];
        // В векторе выше, перечислены далеко не все столбцы, что будут в акте (в акте может быть что угодно и при этом повторяться в неизвестном количестве).
        // В PART_1 мы перечислили, то чему хотели задать порядок заранее, но есть столбцы, где мы хотим оставить порядок, который существует в актах.
        // Чтобы продолжить, поделим отсутсвующие столбцы на два вида: соответсвующие форме акта, заданного в качестве шаблона, и те, которые в его форму не вписались.
        // Столбцы, которые будут совпадать со структурой шаблонного акта, получат приоритет и будут стремится в левое положение таблицы выстраиваясь в том же порядке что и в шаблоне.
        // Другими словами, структура нашего отчета воспроизведет в столбцах порядок итогов из шаблонного акта. Все что не вписальось в эту структуру будет размещено в крайних правых столбцах Excel.
        // В итогах присутсвует два вида данных: базовые и текущие цены, таким образом получается отчет будет написан из 3 частей.

        let part_main = PrintPart::new(main_list).unwrap(); //unwrap не требует обработки: нет идей как это обрабатывать

        Report {
            book: wb,
            part_main,
            part_base: None,
            part_curr: None,
            skip_row: 1,
            body_size: 0,
        }
    }

    pub fn write(&mut self, act: &'a Act) -> Result<(), String> {
        if self.part_base.is_none() && self.part_curr.is_none() {
            let (vec_base, vec_curr) = Self::other_print_parts(act, &self.part_main.vector);
            let part_base = PrintPart::new(vec_base);
            let part_curr = PrintPart::new(vec_curr);

            self.part_base = part_base;
            self.part_curr = part_curr;
        }

        Self::write_header(self, act)?;
        for totalsrow in act.data_of_totals.iter() {
            Self::write_totals(self, totalsrow)?;
        }
        self.body_size += 1;
        Ok(())
    }
    fn write_header(&mut self, act: &Act) -> Result<(), String> {
        let fmt_num = self
            .book
            .add_format()
            .set_num_format(r#"#,##0.00____;-#,##0.00____;"-"____"#);

        let fmt_url = self
            .book
            .add_format()
            .set_font_color(xlsxwriter::FormatColor::Blue)
            .set_underline(xlsxwriter::FormatUnderline::Single);

        let fmt_date = self.book.add_format().set_num_format("dd/mm/yyyy");

        let mut wrapped_sheet = self.book.get_worksheet("Result");

        if wrapped_sheet.is_none() {
            wrapped_sheet = self.book.add_worksheet(Some("Result")).ok();
        };

        let mut sh = wrapped_sheet.unwrap(); //_or(
        let row = self.skip_row + self.body_size + 1;

        let mut column = 0_u16;
        for item in self.part_main.vector.iter() {
            if item.moving == Moving::Del {
                continue;
            }
            if let Source::InTableHeader(name) = item.source {
                let index = act
                    .names_of_header
                    .iter()
                    .position(|desired_data| desired_data.name == name)
                    .unwrap(); //.unwrap_or(return Err(format!("Ошибка в логике программы, сообщающая о необходимости исправления программного кода: \"{}\" обязательно должен быть перечислен в DESIRED_DATA_ARRAY", name)));
                let datavariant = &act.data_of_header[index];

                let date_list = [
                    "Договор дата",
                    "Акт дата",
                    "Отчетный период начало",
                    "Отчетный период окончание",
                ];
                let name_is_date = date_list.contains(&name);
                let format = if name_is_date { Some(&fmt_date) } else { Some(&fmt_num) };

                match datavariant {
                    Some(DataVariant::String(text)) if name_is_date => {
                        let re = Regex::new(r"^\d{2}.\d{2}.\d{4}$").unwrap();
                        if re.is_match(text) {
                            let mut date_iterator =
                                text.split('.').flat_map(|s| s.parse::<i16>().ok()); //.map(|d| d);
                            let day = date_iterator.next().unwrap() as i8;
                            let month = date_iterator.next().unwrap() as i8;
                            let year = date_iterator.next().unwrap();
                            let datetime = DateTime::new(year, month, day, 0, 0, 0.0);

                            sh.write_datetime(row, column, &datetime, format);
                        }
                    }
                    Some(DataVariant::String(text)) => {
                        write_string(&mut sh, row, column, text, None)?
                    }
                    Some(DataVariant::Float(number)) => {
                        write_number(&mut sh, row, column, *number, format)?
                    }
                    None => (),
                }
            }
            if let Source::Calculate(name) = item.source {
                match name {
            "Глава" => loop {
                let index_1 = act.names_of_header.iter().position(|desired_data| desired_data.name == "Глава").unwrap();//_or(return Err("Ошибка в логике программы, сообщающая о необходимости исправления программного кода: \"Глава\" обязательно должна быть в DESIRED_DATA_ARRAY".to_owned()));
                let index_2 = act.names_of_header.iter().position(|desired_data| desired_data.name == "Глава наименование").unwrap();//_or(return Err("Ошибка в логике программы, сообщающая о необходимости исправления программного кода: \"Глава наименование\" обязательно должна быть в DESIRED_DATA_ARRAY".to_owned()));
                let datavariant_1 = &act.data_of_header[index_1];
                let datavariant_2 = &act.data_of_header[index_2];

                let temp_res_1 = match datavariant_1 {
                    Some(DataVariant::String(word)) if !word.is_empty() => word,
                    _ => break,
                };

                let temp_res_2 = match datavariant_2 {
                    Some(DataVariant::String(word)) if !word.is_empty() => word,
                    _ => break,
                };

                let text = format!("{} «{}»", temp_res_1, temp_res_2);
                write_string(&mut sh, row, column, &text, None)?;
                break;
            },
            "Смета №" => {
                let index = act.names_of_header.iter().position(|desired_data| desired_data.name == name).unwrap();//_or(return Err(format!("Ошибка в логике программы, сообщающая о необходимости исправления программного кода: \"{}\" обязательно должен быть перечислен в DESIRED_DATA_ARRAY", name)));
                let datavariant = &act.data_of_header[index];

                if let Some(DataVariant::String(txt)) = datavariant {
                    let text = txt.trim_start_matches("Смета № ");
                        write_string(&mut sh, row, column, text, None);
                }
            }
            "Акт №" => {
                let index = act.names_of_header.iter().position(|desired_data| desired_data.name == name).unwrap();//_or(return Err(format!("Ошибка в логике программы, сообщающая о необходимости исправления программного кода: \"{}\" обязательно должен быть перечислен в DESIRED_DATA_ARRAY", name)));
                let datavariant = &act.data_of_header[index];

                if let Some(DataVariant::String(text)) = datavariant {
                    if text.matches(['/']).count() == 3 {
                       let text = &text.chars().take_while(|ch| *ch != '/').collect::<String>();
                       write_string(&mut sh, row, column, text, None)?;
                    }
                }
            }
            "По смете в ц.2000г." | "Выполнение работ в ц.2000г." => {
                let index = act.names_of_header.iter().position(|desired_data| desired_data.name == name).unwrap();//_or(return Err(format!("Ошибка в логике программы, сообщающая о необходимости исправления программного кода: \"{}\" обязательно должен быть перечислен в DESIRED_DATA_ARRAY", name)));
                let datavariant = &act.data_of_header[index];

                if let Some(DataVariant::String(text)) = datavariant {
                    let _ = text.replace("тыс.", "")
                        .replace("руб.", "")
                        .replace(',', ".")
                        .replace(' ', "")
                        .parse::<f64>()
                        .map(|number| write_number(&mut sh, row, column, number * 1000., Some(&fmt_num))).unwrap();
                }
            }
            "Папка (ссылка)" => {
                if let Some(file_name) = act.path.split('\\').last() {
                    let folder_path = act.path.replace(file_name, "");
                    sh.write_url(row, column, &folder_path, None);
                };
            },
            "Файл (ссылка)" => {
                if let Some(file_name) = act.path.split('\\').last() {
                    let formula = format!("=HYPERLINK(\"{}\", \"{}\")", act.path, file_name);
                    write_formula(&mut sh, row, column, &formula, Some(&fmt_url))?;
                };
            }
            _ => return Err(format!("Ошибка в логике программы, сообщающая о необходимости исправления программного кода: невозможная попытка записать \"{}\" на лист Excel", name)),
        }
            }
            column += item.expected_columns;
        }
        Ok(())
    }

    fn write_totals(&mut self, totalsrow: &TotalsRow) -> Result<(), String> {
        let mut wrapped_sheet = self.book.get_worksheet("Result");

        if wrapped_sheet.is_none() {
            wrapped_sheet = self.book.add_worksheet(Some("Result")).ok();
        };

        let mut sh = wrapped_sheet.unwrap(); //_or(

        let part_main = &self.part_main;
        let part_base = self.part_base.as_ref().unwrap();
        let part_curr = self.part_curr.as_ref().unwrap();
        let row = self.skip_row + self.body_size + 1;

        let get_part = |kind: &str, name: &str| {
            let (part, column_information, corr) = match kind {
                "base" => (
                    "part_base",
                    part_base.get_index_and_address_by_columns(kind, name, Matches::Exact),
                    0,
                ),
                "curr" => (
                    "part_curr",
                    part_curr.get_index_and_address_by_columns(kind, name, Matches::Exact),
                    part_base.get_number_of_columns(),
                ),
                _ => unreachable!("операция не над итоговыми строками акта"),
            };

            match column_information {
                Some((index, col_number_in_vec)) => Some((
                    part,
                    corr + part_main.get_number_of_columns(),
                    index,
                    col_number_in_vec,
                )),
                _ => match part_main.get_index_and_address_by_columns(kind, name, Matches::Exact) {
                    Some((index, col_number_in_vec)) => Some((
                        "part_main",
                        corr + part_main.get_number_of_columns(),
                        index,
                        col_number_in_vec,
                    )),
                    _ => None,
                },
            }
        };

        let fmt_num = self
            .book
            .add_format()
            .set_num_format(r#"#,##0.00____;-#,##0.00____;"-"____"#);

        let mut write_if_some =
            |kind: &str, column_info: Option<(&str, u16, usize, u16)>| -> Result<(), String> {
                if let Some((part, corr, index, col_number_in_vec)) = column_info {
                    let (totalsrow_vec, part) = match part {
                        "part_base" if kind == "base" => (&totalsrow.base_price, part_base),
                        "part_curr" if kind == "curr" => (&totalsrow.curr_price, part_curr),
                        "part_main" if kind == "base" => (&totalsrow.base_price, part_main),
                        "part_main" if kind == "curr" => (&totalsrow.curr_price, part_main),
                        _ => {
                            unreachable!("операция не над итоговыми строками акта")
                        }
                    };

                    let min_number_of_col =
                        (part.vector[index].expected_columns as usize).min(totalsrow_vec.len());
                    for (number_of_col, number) in
                        totalsrow_vec.iter().enumerate().take(min_number_of_col)
                    {
                        if let Some(number) = number {
                            write_number(
                                &mut sh,
                                row,
                                col_number_in_vec + corr + number_of_col as u16,
                                *number,
                                Some(&fmt_num),
                            )?;
                        }
                    }
                }
                Ok(())
            };

        let write_base = get_part("base", &totalsrow.name);
        let write_curr = get_part("curr", &totalsrow.name);
        write_if_some("base", write_base)?;
        write_if_some("curr", write_curr)?;
        Ok(())
    }

    fn other_print_parts(
        sample: &'a Act,
        part_1: &[OutputData],
    ) -> (Vec<OutputData>, Vec<OutputData>) {
        let exclude_from_base = part_1
            .iter()
            .filter(|outputdata| matches!(outputdata.source, Source::AtBasePrices(_, _)))
            .collect::<Vec<_>>();

        let exclude_from_curr = part_1
            .iter()
            .filter(|outputdata| matches!(outputdata.source, Source::AtCurrPrices(_, _)))
            .collect::<Vec<_>>();

        let get_outputdata = |exclude: &[&OutputData],
                              totalsrow: &'a TotalsRow,
                              kind: &str|
         -> Option<OutputData> {
            let name = &totalsrow.name;

            let mut not_listed = true;
            let mut required = false;
            let mut rename = None;
            let matches = Matches::Exact;

            for item in exclude.iter() {
                match item {
                    OutputData {
                        rename: set_name,
                        moving: mov,
                        source:
                            Source::AtBasePrices(text, Matches::Exact)
                            | Source::AtCurrPrices(text, Matches::Exact),
                        ..
                    } if text == name => {
                        not_listed = false;
                        if mov == &Moving::No {
                            required = true;
                            rename = *set_name;

                            println!(
                                "other_print_parts: разрешил '{}' по точному совпадению имени, список досматриваться не будет",
                                name
                            );
                        }
                        break;
                    }
                    OutputData {
                        rename: set_name,
                        moving: mov,
                        source:
                            Source::AtBasePrices(text, Matches::Contains)
                            | Source::AtCurrPrices(text, Matches::Contains),
                        ..
                    } if name.contains(text) => {
                        not_listed = false;
                        if mov == &Moving::No {
                            required = true;
                            rename = *set_name;

                            println!(
                                "other_print_parts: разрешил '{}' по НЕточному совпадению имени, список досматриваться не будет",
                                name
                            );
                        }
                        break;
                    }
                    _ => (),
                }
            }

            if required || not_listed {
                let moving = Moving::No;
                let expected_columns = match kind {
                    "base" => totalsrow.base_price.len() as u16,
                    "curr" => totalsrow.curr_price.len() as u16,
                    _ => {
                        unreachable!("операция не над итоговыми строками акта")
                    }
                };

                let source = match kind {
                    "base" => Source::AtBasePrices(totalsrow.name.clone(), matches),
                    "curr" => Source::AtCurrPrices(totalsrow.name.clone(), matches),
                    _ => {
                        unreachable!("операция не над итоговыми строками акта")
                    }
                };

                let outputdata = OutputData {
                    rename,
                    moving,
                    expected_columns,
                    source,
                };

                return Some(outputdata);
            }
            None
        };

        let (part_base, part_curr) = sample.data_of_totals.iter().fold(
            (Vec::<OutputData>::new(), Vec::<OutputData>::new()),
            |mut acc, smpl_totalsrow| {
                if let Some(x) = get_outputdata(&exclude_from_base, smpl_totalsrow, "base") {
                    acc.0.push(x)
                };

                if let Some(y) = get_outputdata(&exclude_from_curr, smpl_totalsrow, "curr") {
                    acc.1.push(y)
                };
                acc
            },
        );
        (part_base, part_curr)
    }
    pub fn end(self) -> Result<Workbook, String> {
        let mut sh = self
            .book
            .get_worksheet("Result")
            .ok_or("Не удается записать результат в файл Excel")?;

        let header_name: Vec<&OutputData> = self
            .part_main
            .vector
            .iter()
            .filter(|outputdata| {
                outputdata.moving != Moving::Del
                    && !(outputdata.moving == Moving::No
                        && (matches!(outputdata.source, Source::AtBasePrices(_, _))
                            || matches!(outputdata.source, Source::AtCurrPrices(_, _))))
            })
            .chain(self.part_base.as_ref().unwrap().vector.iter())
            .chain(self.part_curr.as_ref().unwrap().vector.iter())
            .collect();

        let header_row = self.skip_row;

        // Это формат для заголовка excel-таблицы
        let fmt_header = self
            .book
            .add_format()
            .set_bold()
            .set_text_wrap() // перенос строк внутри ячейки
            .set_align(FormatAlignment::VerticalTop)
            .set_align(FormatAlignment::Center)
            .set_border(xlsxwriter::FormatBorder::Thin);

        let fmt_first_row_num = self
            .book
            .add_format()
            .set_bold()
            .set_align(FormatAlignment::VerticalTop)
            .set_shrink() // автоуменьшение шрифта текста, если не влез в ячейку
            .set_border(xlsxwriter::FormatBorder::Thin)
            .set_font_size(12.)
            .set_font_color(xlsxwriter::FormatColor::Custom(2050429))
            .set_num_format(r#"#,##0.00____;-#,##0.00____;"-"____"#);

        let fmt_first_row_str = self
            .book
            .add_format()
            .set_bold()
            .set_align(FormatAlignment::VerticalTop)
            .set_align(FormatAlignment::Center)
            .set_shrink() // автоуменьшение шрифта текста, если не влез в ячейку
            .set_border(xlsxwriter::FormatBorder::Thin)
            .set_font_size(12.)
            .set_font_color(xlsxwriter::FormatColor::Custom(2050429));

        let mut counter = 0;
        for outputdata in header_name.iter() {
            let prefix = match outputdata.source {
                Source::AtBasePrices(_, _) => Some("БЦ"),
                Source::AtCurrPrices(_, _) => Some("TЦ"),
                _ => None,
            };

            let ending = match outputdata.rename {
                Some(x) => x.to_owned(),
                _ => match &outputdata.source {
                    Source::InTableHeader(x) => x,
                    Source::Calculate(x) => x,
                    Source::AtBasePrices(x, _) => &x[..],
                    Source::AtCurrPrices(x, _) => &x[..],
                }
                .to_owned(),
            };

            let name = if let Some(x) = prefix {
                x.to_owned() + " " + &ending
            } else {
                ending
            };

            for exp_col in 0..outputdata.expected_columns {
                let col = counter + exp_col;
                write_string(&mut sh, header_row, col, &name, Some(&fmt_header))?;

                //вычисляется ширина столбца excel при переносе строк в ячейке
                let width = if counter < self.part_main.number_of_columns {
                    let name_len = name.chars().count() / 2;
                    let mut first_line_len = 0;
                    for word in name.split(' ') {
                        first_line_len += word.chars().count() + 1;
                        if first_line_len > name_len {
                            break;
                        }
                    }
                    11.max(first_line_len) as f64
                } else {
                    20.11
                };

                sh.set_column(col, col, width, None);
            }
            counter += outputdata.expected_columns;
        }

        let last_row = self.skip_row + self.body_size;
        let last_col = self.part_main.get_number_of_columns()
            + self.part_base.as_ref().unwrap().get_number_of_columns()
            + self.part_curr.as_ref().unwrap().get_number_of_columns()
            - 1;

        let first_row_tab_body = header_row + 1;
        sh.set_row(header_row - 1, 29., None);
        sh.set_row(header_row, 43., None);
        sh.autofilter(header_row, 0, last_row, last_col);

        sh.freeze_panes(first_row_tab_body, 0);

        let (_, column_sbt_103) = self
            .part_main
            .get_index_and_address_by_columns("calc", "Акт №", Matches::Exact)
            .unwrap();
        let col_prefix = column_written_with_letters(column_sbt_103);
        let formula = format!(
            "=SUBTOTAL(103,{col_prefix}{start}:{col_prefix}{end})&\" шт.\"",
            start = first_row_tab_body + 1,
            end = last_row + 1
        );
        write_formula(
            &mut sh,
            header_row - 1,
            column_sbt_103,
            &formula,
            Some(&fmt_first_row_str),
        )?;

        // Вставка формул с подсчетом сумм по excel-таблице
        for column_sbt_109 in self.part_main.get_number_of_columns()..=last_col {
            let col_prefix = column_written_with_letters(column_sbt_109);
            let formula = format!(
                "=SUBTOTAL(109,{col_prefix}{start}:{col_prefix}{end})",
                start = first_row_tab_body + 1,
                end = last_row + 1
            );
            write_formula(
                &mut sh,
                header_row - 1,
                column_sbt_109,
                &formula,
                Some(&fmt_first_row_num),
            )?;
        }
        // let x = sh.set_column_opt(0, 30, 10, format, Some(RowColOptions));

        Ok(self.book)
    }
}

fn write_string(
    sheet: &mut Worksheet,
    row: u32,
    col: u16,
    text: &str,
    format: Option<&Format>,
) -> Result<(), String> {
    sheet.write_string(row, col, text, format).unwrap();
    // _or(
    //     return Err(format!(
    //         "Ошибка записи` строкового значения: \"{}\" в книге Excel",
    //         text
    //     )),
    // );
    Ok(())
}

fn write_number(
    sheet: &mut Worksheet,
    row: u32,
    col: u16,
    number: f64,
    format: Option<&Format>,
) -> Result<(), String> {
    sheet.write_number(row, col, number, format).unwrap();
    // _or(
    //     return Err(format!(
    //         "Ошибка записи` числового значения: \"{}\" в книге Excel",
    //         number
    //     )),
    // );
    Ok(())
}
fn write_formula(
    sheet: &mut Worksheet,
    row: u32,
    col: u16,
    formula: &str,
    format: Option<&Format>,
) -> Result<(), String> {
    sheet.write_formula(row, col, formula, format).unwrap();
    // _or(
    //     return Err(format!(
    //         "Ошибка записи` формулы: \"{}\" в книге Excel",
    //         formula
    //     )),
    // );
    Ok(())
}
fn column_written_with_letters(column: u16) -> String {
    let integer = column / 26;
    let remainder = (column % 26) as u8;
    let ch = char::from(remainder + 65).to_ascii_uppercase().to_string();

    if integer == 0 {
        return ch;
    }

    column_written_with_letters(integer - 1) + &ch
}

#[cfg(test)]
mod tests {
    #[test]
    fn column_in_excel_with_letters_01() {
        use super::column_written_with_letters;
        let result = column_written_with_letters(886);
        assert_eq!(result, "AHC".to_string());
    }
    #[test]
    fn column_in_excel_with_letters_02() {
        use super::column_written_with_letters;
        let result = column_written_with_letters(1465);
        assert_eq!(result, "BDJ".to_string());
    }
}
