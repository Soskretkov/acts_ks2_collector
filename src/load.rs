use crate::transform::{Act, DataVariant, TotalsRow};
use acts_ks2_etl::variant_eq;
use xlsxwriter::{Format, Workbook, Worksheet};

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
    total_col: u16,
}

impl<'a> PrintPart {
    pub fn new(vector: Vec<OutputData>) -> Option<PrintPart> {
        let total_col = Self::count_col(&vector);

        Some(PrintPart { vector, total_col })
    }
    pub fn get_number_of_columns(&self) -> u16 {
        self.total_col
    }

    pub fn get_column(&self, kind: &str, name: &str, matches: Matches) -> Option<(usize, u16)> {
        let src = match kind {
            "base" => Source::AtBasePrices("".to_string(), matches.clone()),
            "curr" => Source::AtCurrPrices("".to_string(), matches.clone()),
            _ => unreachable!("операция не над итоговыми строками акта"),
        };

        let mut counter = 0;
        let mut index = 0;
        for outputdata in self.vector.iter() {
            // println!("{counter}: {:?}", outputdata.source);
            match outputdata {
                OutputData {
                    source: Source::AtBasePrices(text, m) | Source::AtCurrPrices(text, m),
                    ..
                } if acts_ks2_etl::variant_eq(&outputdata.source, &src)
                    && variant_eq(m, &matches)
                    && m == &Matches::Exact
                    && name == text =>
                {
                    return Some((index, counter));
                }
                OutputData {
                    source: Source::AtBasePrices(text, m) | Source::AtCurrPrices(text, m),
                    ..
                } if acts_ks2_etl::variant_eq(&outputdata.source, &src)
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
                _ => (),
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
        printpart.get_column(
            "curr",
            "Стоимость материальных ресурсов (всего)",
            Matches::Exact
        )
    );
    assert_eq!(
        Some((4, 10)),
        printpart.get_column("curr", "Накладные расходы", Matches::Contains)
    );
}
pub struct Report {
    pub book: Option<xlsxwriter::Workbook>,
    pub part_main: PrintPart,
    pub part_base: Option<PrintPart>,
    pub part_curr: Option<PrintPart>,
    pub empty_row: u32,
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
            OutputData{rename: None,                                        moving: Moving::Del, expected_columns: 1, source: Source::Calculate("Ссылка на папку")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::Calculate("Ссылка на файл")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::InTableHeader("Исполнитель")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::Calculate("Глава")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::InTableHeader("Объект")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::InTableHeader("Договор №")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::InTableHeader("Договор дата")},
            // OutputData{rename: None,                                            moving: Moving::Yes,   expected_columns: 1, source: Source::AtBasePrices("Стоимость материальных ресурсов (всего)", Matches::Exact)},
            // OutputData{rename: Some("Восстание машин"),                         moving: Moving::No, expected_columns: 1, source: Source::AtBasePrices("Эксплуатация машин", Matches::Exact)},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::Calculate("Смета №")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::InTableHeader("Смета наименование")},
            OutputData{rename: Some("По смете в ц.2000г., руб."),           moving: Moving::No, expected_columns: 1, source: Source::Calculate("По смете в ц.2000г.")},
            OutputData{rename: Some("Выполнение работ в ц.2000г., руб."),   moving: Moving::No, expected_columns: 1, source: Source::Calculate("Выполнение работ в ц.2000г.")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::Calculate("Акт №")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::InTableHeader("Акт дата")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::InTableHeader("Отчетный период начало")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::InTableHeader("Отчетный период окончание")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::InTableHeader("Метод расчета")},
            OutputData{rename: None,                                        moving: Moving::Del, expected_columns: 1, source: Source::AtBasePrices("Всего с НР и СП (тек".to_string(), Matches::Contains)},
            OutputData{rename: None,                                        moving: Moving::Del, expected_columns: 1, source: Source::AtCurrPrices("Всего с НР и СП (баз".to_string(), Matches::Contains)},
            // OutputData{rename: Some("РЕНЕЙМ................"),                  moving: Moving::No, expected_columns: 1, source: Source::AtBasePrices("Производство работ в зимнее время 4%", Matches::Exact)},
            // OutputData{rename: Some("УДАЛИТЬ..............."),                  moving: Moving::Del, expected_columns: 1, source: Source::AtBasePrices("Итого с К = 1", Matches::Exact)},
            // OutputData{rename: None,                                            moving: Moving::Yes, expected_columns: 1, source: Source::AtBasePrices("ы", Matches::Contains)},
        ];
        // В векторе выше, перечислены далеко не все столбцы, что будут в акте (в акте может быть что угодно и при этом повторяться в неизвестном количестве).
        // В PART_1 мы перечислили, то чему хотели задать порядок заранее, но есть столбцы, где мы хотим оставить порядок, который существует в актах.
        // Чтобы продолжить, поделим отсутсвующие столбцы на два вида: соответсвующие форме акта, заданного в качестве шаблона, и те, которые в его форму не вписались.
        // Столбцы, которые будут совпадать со структурой шаблонного акта, получат приоритет и будут стремится в левое положение таблицы выстраиваясь в том же порядке что и в шаблоне.
        // Другими словами, структура нашего отчета воспроизведет в столбцах порядок итогов из шаблонного акта. Все что не вписальось в эту структуру будет размещено в крайних правых столбцах Excel.
        // В итогах присутсвует два вида данных: базовые и текущие цены, таким образом получается отчет будет написан из 3 частей.
        // let vec_1: Vec<OutputData> = list
        //     .into_iter()
        //     .filter(|outputdata| {
        //         outputdata.moving != Moving::Del
        //             && !(outputdata.moving == Moving::No
        //                 && (matches!(outputdata.source, Source::AtBasePrices(_, _))
        //                     || matches!(outputdata.source, Source::AtCurrPrices(_, _))))
        //     })
        //     .collect();
        let part_main = PrintPart::new(main_list).unwrap(); //unwrap не требует обработки: нет идей как это обрабатывать

        Report {
            book: Some(wb),
            part_main,
            part_base: None,
            part_curr: None,
            empty_row: 1,
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

        let mut wrapped_sheet = self
            .book
            .as_ref()
            .unwrap() //_or(return Err ("Не удалось получить доступ к книге Excel, хранящейся в поле структуры \"Report\"".to_string()))
            .get_worksheet("Result");

        if wrapped_sheet.is_none() {
            wrapped_sheet = self
                .book
                .as_mut()
                .unwrap() //unwrap не требует обработки: обработка в переменной "wrapped_sheet"
                .add_worksheet(Some("Result"))
                .ok();
        };

        let mut sh = wrapped_sheet.unwrap(); //_or(
                                             //     return Err("Не удалось создать лист для отчетной формы внутри книги Excel".to_owned()),
                                             // );

        let mut column = 0_u16;
        // ниже итерация через "for" т.к. обработка ошибок со знаком "?" отклоняет калькулирующие замыкания итератеров таких как "fold"
        for item in self.part_main.vector.iter() {
            match item.moving {
                Moving::Del => continue,
                _ => (),
            }

            if let Source::InTableHeader(name) = item.source {
                Self::write_header(act, name, &mut sh, self.empty_row, column)?;
            }

            if let Source::Calculate(name) = item.source {
                Self::write_calculated(act, name, &mut sh, self.empty_row, column)?;
            }
            column += item.expected_columns;
        }

        for totalsrow in act.data_of_totals.iter() {
            Self::write_totals(
                totalsrow,
                &self.part_main,
                self.part_base.as_ref().unwrap(),
                self.part_curr.as_ref().unwrap(),
                &mut sh,
                self.empty_row,
            );
        }
        self.empty_row += 1;
        Ok(())
    }
    fn write_header(
        act: &Act,
        name: &str,
        sh: &mut Worksheet,
        row: u32,
        col: u16,
    ) -> Result<(), String> {
        let index = act
            .names_of_header
            .iter()
            .position(|desired_data| desired_data.name == name)
            .unwrap(); //.unwrap_or(return Err(format!("Ошибка в логике программы, сообщающая о необходимости исправления программного кода: \"{}\" обязательно должен быть перечислен в DESIRED_DATA_ARRAY", name)));
        let datavariant = &act.data_of_header[index];

        if let Some(DataVariant::String(text)) = datavariant {
            write_string(sh, row, col, text, None)?
        }
        if let Some(DataVariant::Float(number)) = datavariant {
            write_number(sh, row, col, *number, None)?
        }

        Ok(())
    }

    fn write_calculated(
        act: &Act,
        name: &str,
        sh: &mut Worksheet,
        row: u32,
        col: u16,
    ) -> Result<(), String> {
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
                write_string(sh, row, col, &text, None)?;
                break;
            },
            "Смета №" => {
                let index = act.names_of_header.iter().position(|desired_data| desired_data.name == name).unwrap();//_or(return Err(format!("Ошибка в логике программы, сообщающая о необходимости исправления программного кода: \"{}\" обязательно должен быть перечислен в DESIRED_DATA_ARRAY", name)));
                let datavariant = &act.data_of_header[index];

                if let Some(DataVariant::String(text)) = datavariant {
                    text.strip_prefix("Смета № ")
                        .map(|text| write_string(sh, row, col, text, None));
                }
            }
            "Акт №" => {
                let index = act.names_of_header.iter().position(|desired_data| desired_data.name == name).unwrap();//_or(return Err(format!("Ошибка в логике программы, сообщающая о необходимости исправления программного кода: \"{}\" обязательно должен быть перечислен в DESIRED_DATA_ARRAY", name)));
                let datavariant = &act.data_of_header[index];

                if let Some(DataVariant::String(text)) = datavariant {
                    if text.matches(['/']).count() == 3 {
                       let text = &text.chars().take_while(|ch| *ch != '/').collect::<String>();
                       write_string(sh, row, col, text, None);
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
                        .map(|number| write_number(sh, row, col, number * 1000., None)).unwrap();
                }
            }
            "Ссылка на папку" => {},
            "Ссылка на файл" => {
                if let Some(file_name) = act.path.split('\\').last() {
                    let formula = format!("=HYPERLINK(\"{}\", \"{}\")", act.path, file_name);
                    write_formula(sh, row, col, &formula, None)?;
                };
            }
            _ => return Err(format!("Ошибка в логике программы, сообщающая о необходимости исправления программного кода: невозможная попытка записать \"{}\" на лист Excel", name)),
        }

        Ok(())
    }

    fn write_totals(
        totalsrow: &TotalsRow,
        part_main: &PrintPart,
        part_base: &PrintPart,
        part_curr: &PrintPart,
        sh: &mut Worksheet,
        row: u32,
    ) -> Result<(), String> {
        let get_part = |kind: &str, name: &str| {
            let (part, column_information, corr) = match kind {
                "base" => ("base", part_base.get_column(kind, name, Matches::Exact), 0),
                "curr" => (
                    "curr",
                    part_curr.get_column(kind, name, Matches::Exact),
                    part_base.total_col,
                ),
                _ => unreachable!("операция не над итоговыми строками акта"),
            };

            match column_information {
                Some((index, col_number_in_vec)) => {
                    return Some((part, corr + part_main.total_col, index, col_number_in_vec))
                }
                _ => match part_main.get_column(kind, name, Matches::Exact) {
                    Some((index, col_number_in_vec)) => {
                        return Some(("main", corr + part_main.total_col, index, col_number_in_vec))
                    }
                    _ => None,
                },
            }
        };

        let mut write_if_some =
            |column_info: Option<(&str, u16, usize, u16)>| -> Result<(), String> {
                if let Some((part, corr, index, col_number_in_vec)) = column_info {
                    let (totalsrow_vec, part) = match part {
                        "base" => (&totalsrow.base_price, part_base),
                        "curr" => (&totalsrow.curr_price, part_curr),
                        _ => unreachable!("операция не над итоговыми строками акта"),
                    };
                    let min_number_of_col =
                        (part.vector[index].expected_columns as usize).min(totalsrow_vec.len());
                    for number_of_col in 0..min_number_of_col {
                        let number = totalsrow_vec[number_of_col];
                        if let Some(number) = number {
                            write_number(
                                sh,
                                row,
                                col_number_in_vec + corr + number_of_col as u16,
                                number,
                                None,
                            )?;
                        }
                    }
                }
                Ok(())
            };

        let write_base = get_part("base", &totalsrow.name);
        let write_curr = get_part("curr", &totalsrow.name);
        write_if_some(write_base)?;
        write_if_some(write_curr)?;
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
                if let Some(x) = get_outputdata(&exclude_from_base, &smpl_totalsrow, "base") {
                    acc.0.push(x)
                };

                if let Some(y) = get_outputdata(&exclude_from_curr, &smpl_totalsrow, "curr") {
                    acc.1.push(y)
                };
                acc
            },
        );
        (part_base, part_curr)
    }
    pub fn end(&mut self) -> Option<Workbook> {
        let mut sh = self
            .book
            .as_ref()
            .unwrap() //_or(return Err ("Не удалось получить доступ к книге Excel, хранящейся в поле структуры \"Report\"".to_string()))
            .get_worksheet("Result")
            .unwrap();

        let first_row = self
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
            .chain(self.part_curr.as_ref().unwrap().vector.iter());

        first_row.fold(0, |mut acc, outputdata| {
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

            let name = if prefix.is_some() {
                prefix.unwrap().to_owned() + " " + &ending
            } else {
                ending
            };

            (0..outputdata.expected_columns).for_each(|exp_col| {
                write_string(&mut sh, 0, acc + exp_col, &name, None);
            });
            acc + outputdata.expected_columns
        });

        self.book.take()
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
