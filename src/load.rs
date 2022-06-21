use crate::transform::{Act, DataVariant};
use itertools::fold;
use xlsxwriter::{Workbook, Worksheet};

#[derive(Debug, Clone)]
pub struct OutputData<'a> {
    pub rename: Option<&'a str>,
    pub moving: Moving,
    pub expected_columns: u16,
    pub source: Source<'a>,
}
#[derive(Debug, Clone, PartialEq)]

pub enum Moving {
    Already,
    Remain,
    Move,
    Delete,
}

// Четыре вида данных на выходе: в готовом виде в шапке, в готов виде в итогах акта (2 варанта), и нет готовых (нужно расчитать программой):
#[derive(Debug, Clone, PartialEq)]
pub enum Source<'a> {
    InTableHeader(&'static str),
    AtCurrPrices(&'a str),
    AtBasePrices(&'a str),
    Calculate(&'static str),
}

#[derive(Debug)]
pub struct PrintPart<'a> {
    vector: Vec<OutputData<'a>>,
    total_col: u16,
}

impl<'a> PrintPart<'a> {
    pub fn new(vector: Vec<OutputData>) -> PrintPart {
        let total_col = Self::count_col(&vector);

        PrintPart { vector, total_col }
    }
    pub fn get_number_of_columns(&self) -> u16 {
        self.total_col
    }
    
    pub fn get_column_number(part: &mut PrintPart<'a>, source: Source<'a>) {
        let mut left_indent: Option<u16> = None;
        let mut counter_2: u16 = 0;

        for outputdata_2 in part.vector.iter() {
            if matches!(&outputdata_2.source, source) {
                left_indent = Some(counter_2);
                break;
            }
            counter_2 += outputdata_2.expected_columns;
        }

        let mut counter_1: u16 = 0;
        for outputdata_1 in part.vector.iter() {
            if let OutputData {
                moving: Moving::Move,
                source,
                ..
            } = outputdata_1
            {
                left_indent = Some(counter_1);
                break;
            }

            counter_1 += outputdata_1.expected_columns;
        }
        
        // let new_data = OutputData {
        //     rename: None,
        //     moving: Moving::Already,
        //     expected_columns,
        //     source: source,
        // };
        // part.vector.push(new_data);
    }
    




    fn count_col(vector: &[OutputData]) -> u16 {
        vector.iter().fold(0, |acc, copy| match copy.moving {
            Moving::Already => acc + copy.expected_columns,
            Moving::Move => acc + copy.expected_columns,
            _ => acc,
        })
    }
}
#[test]
fn PrintPart_test() {
    #[rustfmt::skip]
        let vec_to_test = vec![
            OutputData{rename: None,                           moving: Moving::Already, expected_columns: 1, source: Source::InTableHeader("Объект")},
            OutputData{rename: None,                           moving: Moving::Move,   expected_columns: 1, source: Source::AtCurrPrices("Стоимость материальных ресурсов (всего)")},
            OutputData{rename: Some("РЕНЕЙМ................"), moving: Moving::Remain, expected_columns: 8, source: Source::AtCurrPrices("Производство работ в зимнее время 4%")},
            OutputData{rename: Some("УДАЛИТЬ..............."), moving: Moving::Delete, expected_columns: 5, source: Source::AtCurrPrices("Производство работ в зимнее время 4%")},
        ];
    let printpart = PrintPart::new(vec_to_test);

    assert_eq!(2, printpart.get_number_of_columns());
}

pub struct Report<'a> {
    pub book: Option<xlsxwriter::Workbook>,
    pub part_1_just: PrintPart<'a>,
    pub part_2_base: PrintPart<'a>,
    pub part_3_curr: PrintPart<'a>,
    pub empty_row: u32,
}

impl<'a> Report<'a> {
    pub fn set_sample(wb: xlsxwriter::Workbook, sample: &'a Act) -> Result<Report, &'static str> {
        // Нужно чтобы код назначал длину таблицы по горизонтали в зависимости от количества строк в итогах (обычно итоги имеют 17 строк,
        // но если какой-то акт имеет 16, 18, 0 или, скажем, 40 строк в итогах, то нужна какая-то логика, чтобы соотнести эти 40 строк одного акта
        // с 17 строками других актов. Нужно решение, как не сокращать эти 40 строк до 17 стандартных и выдать информацию пользователю без потерь.
        // Данные делятся на ожидаемые (им порядок можно сразу задать) и случайные.
        // Ниже массив, содержащий информацию о колонках, которые мы ожидаем получить из актов, здесь будем задавать порядок.
        // Позиция в массиве будет соответсвовать столбцу выходной формы (это крайние левые столбцы шапки):

        #[rustfmt::skip]
        let vec_1 = vec![
            OutputData{rename: None,                                        moving: Moving::Already, expected_columns: 1, source: Source::InTableHeader("Исполнитель")},
            OutputData{rename: None,                                        moving: Moving::Already, expected_columns: 1, source: Source::Calculate("Глава")},
            OutputData{rename: None,                                        moving: Moving::Already, expected_columns: 1, source: Source::InTableHeader("Объект")},
            OutputData{rename: None,                                        moving: Moving::Already, expected_columns: 1, source: Source::InTableHeader("Договор №")},
            OutputData{rename: None,                                        moving: Moving::Already, expected_columns: 1, source: Source::InTableHeader("Договор дата")},
            OutputData{rename: None,                                            moving: Moving::Move,   expected_columns: 1, source: Source::AtBasePrices("Стоимость материальных ресурсов (всего)")},
            OutputData{rename: Some("Восстание машин"),                         moving: Moving::Remain, expected_columns: 1, source: Source::AtBasePrices("Эксплуатация машин")},
            OutputData{rename: None,                                        moving: Moving::Already, expected_columns: 1, source: Source::InTableHeader("Смета №")},
            OutputData{rename: None,                                        moving: Moving::Already, expected_columns: 1, source: Source::InTableHeader("Смета наименование")},
            OutputData{rename: Some("По смете в ц.2000г., руб."),           moving: Moving::Already, expected_columns: 1, source: Source::Calculate("По смете в ц.2000г.")},
            OutputData{rename: Some("Выполнение работ в ц.2000г., руб."),   moving: Moving::Already, expected_columns: 1, source: Source::Calculate("Выполнение работ в ц.2000г.")},
            OutputData{rename: None,                                        moving: Moving::Already, expected_columns: 1, source: Source::InTableHeader("Акт №")},
            OutputData{rename: None,                                        moving: Moving::Already, expected_columns: 1, source: Source::InTableHeader("Акт дата")},
            OutputData{rename: None,                                        moving: Moving::Already, expected_columns: 1, source: Source::InTableHeader("Отчетный период начало")},
            OutputData{rename: None,                                        moving: Moving::Already, expected_columns: 1, source: Source::InTableHeader("Отчетный период окончание")},
            OutputData{rename: None,                                        moving: Moving::Already, expected_columns: 1, source: Source::InTableHeader("Метод расчета")},
            OutputData{rename: None,                                        moving: Moving::Already, expected_columns: 1, source: Source::Calculate("Ссылка на папку")},
            OutputData{rename: None,                                        moving: Moving::Already, expected_columns: 1, source: Source::Calculate("Ссылка на файл")},
            OutputData{rename: Some("РЕНЕЙМ................"),                  moving: Moving::Remain, expected_columns: 1, source: Source::AtBasePrices("Производство работ в зимнее время 4%")},
            OutputData{rename: Some("УДАЛИТЬ..............."),                  moving: Moving::Delete, expected_columns: 1, source: Source::AtBasePrices("Итого с К = 1")},
        ];
        // В векторе выше, перечислены далеко не все столбцы, что будут в акте (в акте может быть что угодно и при этом повторяться в неизвестном количестве).
        // В PART_1 мы перечислили, то чему хотели задать порядок заранее, но есть столбцы, где мы хотим оставить порядок, который существует в актах.
        // Чтобы продолжить, поделим отсутсвующие столбцы на два вида: соответсвующие форме акта, заданного в качестве шаблона, и те, которые в его форму не вписались.
        // Столбцы, которые будут совпадать со структурой шаблонного акта, получат приоритет и будут стремится в левое положение таблицы выстраиваясь в том же порядке что и в шаблоне.
        // Другими словами, структура нашего отчета воспроизведет в столбцах порядок итогов из шаблонного акта. Все что не вписальось в эту структуру будет размещено в крайних правых столбцах Excel.
        // В итогах присутсвует два вида данных: базовые и текущие цены, таким образом получается отчет будет написан из 3 частей.

        let (vec_2, vec_3) = Self::other_parts(sample, &vec_1);

        let part_1_just = PrintPart::new(vec_1);
        let part_2_base = PrintPart::new(vec_2);
        let part_3_curr = PrintPart::new(vec_3);

        Ok(Report {
            book: Some(wb),
            part_1_just,
            part_2_base,
            part_3_curr,
            empty_row: 0,
        })
    }
    fn other_parts(
        sample: &'a Act,
        part_1: &[OutputData<'a>],
    ) -> (Vec<OutputData<'a>>, Vec<OutputData<'a>>) {
        let exclude_from_base = part_1
            .iter()
            .filter(|outputdata| matches!(outputdata.source, Source::AtBasePrices(_)))
            .collect::<Vec<_>>();

        let exclude_from_curr = part_1
            .iter()
            .filter(|outputdata| matches!(outputdata.source, Source::AtCurrPrices(_)))
            .collect::<Vec<_>>();

        let (part_2_base, part_3_curr) = sample.data_of_totals.iter().fold(
            (Vec::<OutputData>::new(), Vec::<OutputData>::new()),
            |mut acc, smpl_totalsrow| {
                let (check_renaming_base, not_listed_base, new_name_base) =
                    exclude_from_base.iter().fold(
                        (false, true, None),
                        |(mut it_remains, mut not_listed, mut new_name), item| {
                            match item {
                                OutputData {
                                    rename: set_name,
                                    moving: Moving::Remain,
                                    source: Source::AtBasePrices(name),
                                    ..
                                } if *name == smpl_totalsrow.name => {
                                    it_remains = true;
                                    not_listed = false;
                                    new_name = *set_name;
                                }
                                OutputData {
                                    source: Source::AtBasePrices(name),
                                    ..
                                } if *name == smpl_totalsrow.name => not_listed = false,
                                _ => (),
                            }

                            (it_remains, not_listed, new_name)
                        },
                    );

                if check_renaming_base || not_listed_base {
                    let columns_min = smpl_totalsrow
                        .base_price
                        .iter()
                        .map(Option::is_some)
                        .count() as u16;

                    let outputdata_base = OutputData {
                        rename: new_name_base,
                        moving: Moving::Already,
                        expected_columns: columns_min,
                        source: Source::AtBasePrices(&smpl_totalsrow.name),
                    };
                    acc.0.push(outputdata_base)
                }

                let (check_renaming_curr, not_listed_curr, new_name_curr) =
                    exclude_from_curr.iter().fold(
                        (false, true, None),
                        |(mut it_remains, mut not_listed, mut new_name), item| {
                            match item {
                                OutputData {
                                    rename: set_name,
                                    moving: Moving::Remain,
                                    source: Source::AtCurrPrices(name),
                                    ..
                                } if *name == smpl_totalsrow.name => {
                                    it_remains = true;
                                    not_listed = false;
                                    new_name = *set_name;
                                }
                                OutputData {
                                    source: Source::AtCurrPrices(name),
                                    ..
                                } if *name == smpl_totalsrow.name => not_listed = false,
                                _ => (),
                            }

                            (it_remains, not_listed, new_name)
                        },
                    );

                if check_renaming_curr || not_listed_curr {
                    let columns_min = smpl_totalsrow
                        .current_price
                        .iter()
                        .map(Option::is_some)
                        .count() as u16;

                    let outputdata_curr = OutputData {
                        rename: new_name_curr,
                        moving: Moving::Already,
                        expected_columns: columns_min,
                        source: Source::AtCurrPrices(&smpl_totalsrow.name),
                    };
                    acc.1.push(outputdata_curr)
                }
                acc
            },
        );
        (part_2_base, part_3_curr)
    }

    pub fn write(&mut self, act: &Act) {
        let mut sh = self
            .book
            .as_mut()
            .unwrap()
            .add_worksheet(Some("Result"))
            .unwrap();

        self.part_1_just
            .vector
            .iter()
            .fold(0_u16, |first_col, item| {
                if let Source::InTableHeader(name) = item.source {
                    Self::write_header(&act, &name, &mut sh, self.empty_row, first_col)
                }

                if let Source::Calculate(name) = item.source {
                    Self::write_calculated(&act, &name, &mut sh, self.empty_row, first_col)
                }
                first_col + item.expected_columns
            });

        for row in act.data_of_totals.iter() {}
    }
    // fn write_totals(part: &mut PrintPart<'a>, source: Source<'a>, expected_columns: u16) {

    fn write_header(act: &Act, name: &str, sh: &mut Worksheet, row: u32, col: u16) {
        let index = act.names_of_header.iter().position(|desired_data| desired_data.name == name).expect(&format!("Ошибка в логике программы, сообщающая о необходимости исправления программного кода: \"{}\" обязательно должен быть перечислен в DESIRED_DATA_ARRAY", name));
        let datavariant = &act.data_of_header[index];

        if let Some(DataVariant::String(insert)) = datavariant {
            sh.write_string(row, col, insert, None).unwrap()
        }

        if let Some(DataVariant::Float(x)) = datavariant {
            sh.write_number(row, col, *x, None).unwrap()
        }
    }

    fn write_calculated(act: &Act, name: &str, sh: &mut Worksheet, row: u32, col: u16) {
        match name {
            "Глава" => loop {
                let index_1 = act.names_of_header.iter().position(|desired_data| desired_data.name == "Глава").expect(&format!("Ошибка в логике программы, сообщающая о необходимости исправления программного кода: \"Глава\" обязательно должна быть в DESIRED_DATA_ARRAY"));
                let index_2 = act.names_of_header.iter().position(|desired_data| desired_data.name == "Глава наименование").expect(&format!("Ошибка в логике программы, сообщающая о необходимости исправления программного кода: \"Глава наименование\" обязательно должна быть в DESIRED_DATA_ARRAY"));
                let datavariant_1 = &act.data_of_header[index_1];
                let datavariant_2 = &act.data_of_header[index_2];

                let temp_res_1 = match datavariant_1 {
                    Some(DataVariant::String(word)) if word.len() > 0 => word,
                    _ => break,
                };

                let temp_res_2 = match datavariant_2 {
                    Some(DataVariant::String(word)) if word.len() > 0 => word,
                    _ => break,
                };

                let temp_res_3 = format!("{} «{}»", temp_res_1, temp_res_2);
                sh.write_string(row, col, &temp_res_3, None).unwrap();
                break;
            },
            "По смете в ц.2000г." | "Выполнение работ в ц.2000г." => {
                let index = act.names_of_header.iter().position(|desired_data| desired_data.name == name).expect(&format!("Ошибка в логике программы, сообщающая о необходимости исправления программного кода: \"{}\" обязательно должен быть перечислен в DESIRED_DATA_ARRAY", name));
                let datavariant = &act.data_of_header[index];

                if let Some(DataVariant::String(text)) = datavariant {
                    let _ = text.replace("тыс.", "")
                        .replace("руб.", "")
                        .replace(",", ".")
                        .replace(" ", "")
                        .parse::<f64>()
                        .map(|insert| sh.write_number(row, col, insert * 1000., None).unwrap());
                }
            }
            "Ссылка на папку" => {},
            "Ссылка на файл" => {
                act.path.split("\\").last().map(|file_name| {
                    let formula = format!("=HYPERLINK(\"{}\", \"{}\")", act.path, file_name);
                    sh.write_formula(row, col, &formula, None).unwrap();
                });
            }
            _ => unreachable!("Ошибка в логике программы, сообщающая о необходимости исправления программного кода: невозможная попытка записать \"{}\" на лист Excel", name),
        }
    }

    pub fn stop_writing(&mut self) -> Option<Workbook> {
        let x = self.book.take();
        x
    }
}
