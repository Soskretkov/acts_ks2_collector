use crate::transform::{Act, DataVariant, TotalsRow};
use xlsxwriter::{Format, Workbook, Worksheet};

#[derive(Debug)]
pub struct OutputData<'a> {
    pub rename: Option<&'static str>,
    pub moving: Moving,
    pub expected_columns: u16,
    pub source: Source<'a>,
}
#[derive(Debug, Clone, PartialEq)]

pub enum Moving {
    No,
    Move,
    Delete,
}

#[derive(Debug, PartialEq)]
pub enum Matches {
    Exact,
    Contains,
}

// Четыре вида данных на выходе: в готовом виде в шапке, в готов виде в итогах акта (2 варанта), и нет готовых (нужно расчитать программой):
#[derive(Debug, PartialEq)]
pub enum Source<'a> {
    InTableHeader(&'static str),
    AtCurrPrices(&'a str, Matches),
    AtBasePrices(&'a str, Matches),
    Calculate(&'static str),
}

#[derive(Debug)]
pub struct PrintPart<'a> {
    vector: Vec<OutputData<'a>>,
    total_col: u16,
}

impl<'a> PrintPart<'a> {
    pub fn new(vector: Vec<OutputData>) -> Option<PrintPart> {
        let total_col = Self::count_col(&vector);

        Some(PrintPart { vector, total_col })
    }
    pub fn get_number_of_columns(&self) -> u16 {
        self.total_col
    }

    pub fn get_column(&self, src: &Source<'a>) -> Option<u16> {
        let name = match src {
            Source::AtBasePrices(words, _) => words,
            Source::AtCurrPrices(words, _) => words,
            _ => unreachable!(
                "функция get_column только для строк, которые встречаются в итогах акта"
            ),
        };

        let mut counter = 0;
        for outputdata in self.vector.iter() {
            match outputdata {
                OutputData {
                    source: Source::AtBasePrices(text, Matches::Exact),
                    ..
                } if text == name => {
                    return Some(counter);
                }
                OutputData {
                    source: Source::AtCurrPrices(text, Matches::Exact),
                    ..
                } if text == name => {
                    return Some(counter);
                }
                OutputData {
                    source: Source::AtBasePrices(text, Matches::Contains),
                    ..
                } if name.contains(text) => {
                    return Some(counter);
                }
                OutputData {
                    source: Source::AtCurrPrices(text, Matches::Contains),
                    ..
                } if name.contains(text) => {
                    return Some(counter);
                }
                OutputData {
                    moving: Moving::No, ..
                } => {
                    counter += outputdata.expected_columns;
                }
                OutputData {
                    moving: Moving::Move,
                    ..
                } => {
                    counter += outputdata.expected_columns;
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
                Moving::Move => acc + outputdata.expected_columns,
                _ => acc,
            })
    }
}
#[test]
fn PrintPart_test() {
    #[rustfmt::skip]
        let vec_to_test = vec![
            OutputData{rename: None,                           moving: Moving::No, expected_columns: 1, source: Source::InTableHeader("Объект")},
            OutputData{rename: None,                           moving: Moving::Move, expected_columns: 2, source: Source::AtCurrPrices("Эксплуатация машин", Matches::Exact)},
            OutputData{rename: None,                           moving: Moving::Move, expected_columns: 3, source: Source::AtBasePrices("Эксплуатация машин", Matches::Exact)},
            OutputData{rename: None,                           moving: Moving::Move, expected_columns: 4, source: Source::AtBasePrices("Накладные расходы", Matches::Contains)},
            OutputData{rename: None,                           moving: Moving::Move, expected_columns: 5, source: Source::AtCurrPrices("Накладные расходы и доходы", Matches::Contains)},
            OutputData{rename: Some("РЕНЕЙМ................"), moving: Moving::No, expected_columns: 6, source: Source::AtCurrPrices("Производство работ в зимнее время 4%", Matches::Exact)},
            OutputData{rename: Some("УДАЛИТЬ..............."), moving: Moving::Delete, expected_columns: 7, source: Source::AtCurrPrices("Производство работ в зимнее время 4%", Matches::Exact)},
            OutputData{rename: None,                           moving: Moving::Move,   expected_columns: 8, source: Source::AtCurrPrices("Стоимость материальных ресурсов (всего)", Matches::Exact)},
        ];
    let printpart = PrintPart::new(vec_to_test).unwrap();

    assert_eq!(&23, &printpart.get_number_of_columns());
    assert_eq!(
        Some(15),
        printpart.get_column(&Source::AtCurrPrices(
            "Стоимость материальных ресурсов (всего)",
            Matches::Exact
        ))
    );
    assert_eq!(
        Some(6),
        printpart.get_column(&Source::AtCurrPrices(
            "Накладные расходы",
            Matches::Contains
        ))
    );
}
pub struct Report<'a> {
    pub book: Option<xlsxwriter::Workbook>,
    pub part_main: PrintPart<'a>,
    pub part_base: Option<PrintPart<'a>>,
    pub part_curr: Option<PrintPart<'a>>,
    pub empty_row: u32,
}

impl<'a> Report<'a> {
    pub fn new(wb: xlsxwriter::Workbook) -> Report<'a> {
        // Нужно чтобы код назначал длину таблицы по горизонтали в зависимости от количества строк в итогах (обычно итоги имеют 17 строк,
        // но если какой-то акт имеет 16, 18, 0 или, скажем, 40 строк в итогах, то нужна какая-то логика, чтобы соотнести эти 40 строк одного акта
        // с 17 строками других актов. Нужно решение, как не сокращать эти 40 строк до 17 стандартных и выдать информацию пользователю без потерь.
        // Данные делятся на ожидаемые (им порядок можно сразу задать) и случайные.
        // Ниже массив, содержащий информацию о колонках, которые мы ожидаем получить из актов, здесь будем задавать порядок.
        // Позиция в массиве будет соответсвовать столбцу выходной формы (это крайние левые столбцы шапки):

        #[rustfmt::skip]
        let list: [OutputData; 20] = [
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::InTableHeader("Исполнитель")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::Calculate("Глава")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::InTableHeader("Объект")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::InTableHeader("Договор №")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::InTableHeader("Договор дата")},
            OutputData{rename: None,                                            moving: Moving::Move,   expected_columns: 1, source: Source::AtBasePrices("Стоимость материальных ресурсов (всего)", Matches::Exact)},
            OutputData{rename: Some("Восстание машин"),                         moving: Moving::No, expected_columns: 1, source: Source::AtBasePrices("Эксплуатация машин", Matches::Exact)},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::Calculate("Смета №")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::InTableHeader("Смета наименование")},
            OutputData{rename: Some("По смете в ц.2000г., руб."),           moving: Moving::No, expected_columns: 1, source: Source::Calculate("По смете в ц.2000г.")},
            OutputData{rename: Some("Выполнение работ в ц.2000г., руб."),   moving: Moving::No, expected_columns: 1, source: Source::Calculate("Выполнение работ в ц.2000г.")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::Calculate("Акт №")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::InTableHeader("Акт дата")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::InTableHeader("Отчетный период начало")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::InTableHeader("Отчетный период окончание")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::InTableHeader("Метод расчета")},
            OutputData{rename: None,                                        moving: Moving::Delete, expected_columns: 1, source: Source::Calculate("Ссылка на папку")},
            OutputData{rename: None,                                        moving: Moving::No, expected_columns: 1, source: Source::Calculate("Ссылка на файл")},
            OutputData{rename: Some("РЕНЕЙМ................"),                  moving: Moving::No, expected_columns: 1, source: Source::AtBasePrices("Производство работ в зимнее время 4%", Matches::Exact)},
            OutputData{rename: Some("УДАЛИТЬ..............."),                  moving: Moving::Delete, expected_columns: 1, source: Source::AtBasePrices("Итого с К = 1", Matches::Exact)},
        ];
        // В векторе выше, перечислены далеко не все столбцы, что будут в акте (в акте может быть что угодно и при этом повторяться в неизвестном количестве).
        // В PART_1 мы перечислили, то чему хотели задать порядок заранее, но есть столбцы, где мы хотим оставить порядок, который существует в актах.
        // Чтобы продолжить, поделим отсутсвующие столбцы на два вида: соответсвующие форме акта, заданного в качестве шаблона, и те, которые в его форму не вписались.
        // Столбцы, которые будут совпадать со структурой шаблонного акта, получат приоритет и будут стремится в левое положение таблицы выстраиваясь в том же порядке что и в шаблоне.
        // Другими словами, структура нашего отчета воспроизведет в столбцах порядок итогов из шаблонного акта. Все что не вписальось в эту структуру будет размещено в крайних правых столбцах Excel.
        // В итогах присутсвует два вида данных: базовые и текущие цены, таким образом получается отчет будет написан из 3 частей.

        let vec_1: Vec<OutputData> = list
            .into_iter()
            // .filter(|outputdata| {
            //     outputdata.moving == Moving::No || outputdata.moving == Moving::Move
            // })
            .collect();
        let part_main = PrintPart::new(vec_1).unwrap(); //unwrap не требует обработки: нет идей как это обрабатывать

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
                Moving::Delete => continue,
                Moving::No => continue,
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
                self.part_base.as_mut().unwrap(),
                self.part_curr.as_mut().unwrap(),
                &mut sh,
                self.empty_row,
            );
        }

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
        part_base: &mut PrintPart,
        part_curr: &mut PrintPart,
        sh: &mut Worksheet,
        row: u32,
    ) {
        let get_column = |source: Source| {
            let column_number_in_totals_vector = match source {
                Source::AtBasePrices(_, _) => part_base.get_column(&source),
                Source::AtCurrPrices(_, _) => match part_curr.get_column(&source) {
                    Some(x) => Some(x + part_base.total_col),
                    None => None,
                },
                _ => unreachable!("операция не над строками, которые встречаются в итогах акта"),
            };

            match column_number_in_totals_vector {
                Some(x) => Some(x + part_main.total_col),
                _ => match part_main.get_column(&source) {
                    Some(x) => Some(x),
                    _ => None,
                },
            }
        };

        if let Some(col) = get_column(Source::AtBasePrices(&totalsrow.name, Matches::Exact)) {
            write_number(sh, row, col, totalsrow.base_price[0].unwrap_or(0.), None);
        } else if let Some(col) =
            get_column(Source::AtBasePrices(&totalsrow.name, Matches::Contains))
        {
            write_number(sh, row, col, totalsrow.base_price[0].unwrap_or(0.), None);
        }

        if let Some(col) = get_column(Source::AtCurrPrices(&totalsrow.name, Matches::Exact)) {
            write_number(sh, row, col, totalsrow.curr_price[0].unwrap_or(0.), None);
        } else if let Some(col) =
            get_column(Source::AtCurrPrices(&totalsrow.name, Matches::Contains))
        {
            write_number(sh, row, col, totalsrow.curr_price[0].unwrap_or(0.), None);
        }
    }

    fn other_print_parts(
        sample: &'a Act,
        part_1: &[OutputData<'a>],
    ) -> (Vec<OutputData<'a>>, Vec<OutputData<'a>>) {
        let exclude_from_base = part_1
            .iter()
            .filter(|outputdata| matches!(outputdata.source, Source::AtBasePrices(_, _)))
            .collect::<Vec<_>>();

        let exclude_from_curr = part_1
            .iter()
            .filter(|outputdata| matches!(outputdata.source, Source::AtCurrPrices(_, _)))
            .collect::<Vec<_>>();

        println!("base {}", exclude_from_base.len());
        println!("curr {}", exclude_from_curr.len());
        let get_outputdata = |exclude: &[&OutputData<'a>],
                              price: &[Option<f64>],
                              source: Source<'a>|
         -> Option<OutputData<'a>> {
            let (name, _) = match &source {
                Source::AtBasePrices(words, matches) => (words, matches),
                Source::AtCurrPrices(words, matches) => (words, matches),
                _ => unreachable!("операция не над строками, которые встречаются в итогах акта"),
            };
            let mut not_listed = true;
            let mut it_another_vector = false;
            let mut new_name = None;
            let mut matches: Matches;

            for item in exclude.iter() {
                match item {
                    OutputData {
                        rename: set_name,
                        moving: Moving::No,
                        source: Source::AtBasePrices(text, m) | Source::AtCurrPrices(text, m),
                        ..
                    } if text == name || (m == &Matches::Contains && name.contains(text)) => {
                        not_listed = false;
                        it_another_vector = true;
                        new_name = *set_name;
                        matches = match *m {
                            Matches::Exact=> Matches::Exact,
                            Matches::Contains => Matches::Contains,
                        };
                        println!("Этот нашелся: {} (теперь ищем следующий)", name);
                        break;
                    }
                    _ => (),
                };
                match item {
                    OutputData {
                        source: Source::AtBasePrices(text, m) | Source::AtCurrPrices(text, m),
                        ..
                    } if not_listed
                        && (text == name || (m == &Matches::Contains && name.contains(text))) =>
                    {
                        not_listed = false;
                        println!("Этот listed: {} (продолжаем смотреть)", name);
                    }
                    _ => (),
                }
            }

            if it_another_vector || not_listed {
                let columns_min = price.iter().map(Option::is_some).count() as u16;

                let outputdata = OutputData {
                    rename: new_name,
                    moving: Moving::No,
                    expected_columns: columns_min,
                    source,
                };

                return Some(outputdata);
            }
            None
        };

        let (part_base, part_curr) = sample.data_of_totals.iter().fold(
            (Vec::<OutputData>::new(), Vec::<OutputData>::new()),
            |mut acc, smpl_totalsrow| {
                if let Some(x) = get_outputdata(
                    &exclude_from_base,
                    &smpl_totalsrow.base_price,
                    Source::AtBasePrices(&smpl_totalsrow.name, Matches::Exact),
                ) {
                    acc.0.push(x)
                };

                // if let Some(y) = get_outputdata(
                //     &exclude_from_curr,
                //     &smpl_totalsrow.curr_price,
                //     Source::AtCurrPrices(&smpl_totalsrow.name, Matches::Exact),
                // ) {
                //     acc.1.push(y)
                // };

                acc
            },
        );
        (part_base, part_curr)
    }
    pub fn finish_writing(&mut self) -> Option<Workbook> {
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
            .chain(self.part_base.as_ref().unwrap().vector.iter())
            .chain(self.part_curr.as_ref().unwrap().vector.iter());

        first_row.fold(0, |mut acc, outputdata| {
            let prefix = match outputdata.source {
                Source::AtBasePrices(_, _) => Some("БЦ"),
                Source::AtCurrPrices(_, _) => Some("TЦ"),
                _ => None,
            };

            let ending = match outputdata.rename {
                Some(x) => x,
                _ => match outputdata.source {
                    Source::InTableHeader(x) => x,
                    Source::Calculate(x) => x,
                    Source::AtBasePrices(x, _) => x,
                    Source::AtCurrPrices(x, _) => x,
                },
            };

            let name = if prefix.is_some() {
                prefix.unwrap().to_owned() + " " + ending
            } else {
                ending.to_owned()
            };

            (0..outputdata.expected_columns).for_each(|exp_col| {
                write_string(&mut sh, 0, acc, &name, None);
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
