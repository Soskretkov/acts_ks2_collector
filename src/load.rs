use crate::transform::Act;

#[derive(Debug, Clone)]
pub struct OutputData<'a> {
    pub set_name: Option<&'a str>,
    pub moving: Moving,
    pub expected_columns: usize,
    pub source: Source<'a>,
}
// Четыре вида данных на выходе: в готовом виде в шапке, в готов виде в итогах акта (2 варанта), и нет готовых (нужно расчитать программой):
#[derive(Debug, Clone, PartialEq)]

pub enum Moving {
    Remain,
    Move,
    Delete,
}

#[derive(Debug, Clone, PartialEq)]
pub enum Source<'a> {
    InTableHeader(&'static str),
    AtCurrPrices(&'a str),
    AtBasePrices(&'a str),
    Calculate,
}

pub struct PrintPart<'a> {
    vector: Vec<OutputData<'a>>,
    total_col: usize,
}

impl<'a> PrintPart<'a> {
    pub fn new(vector: Vec<OutputData>) -> PrintPart {
        let total_col = Self::count_col(&vector);

        PrintPart { vector, total_col }
    }
    pub fn get_number_of_columns(&self) -> usize {
        self.total_col
    }
    fn count_col(vector: &[OutputData]) -> usize {
        vector.iter().fold(0, |acc, copy| match copy.moving {
            Moving::Delete => acc,
            _ => acc + copy.expected_columns,
        })
    }
}
#[test]
fn PrintPart_test() {
    #[rustfmt::skip]
        let vec_to_test = vec![
            OutputData{set_name: None,                           moving: Moving::Remain, expected_columns: 1, source: Source::InTableHeader("Объект")},
            OutputData{set_name: None,                           moving: Moving::Move,   expected_columns: 1, source: Source::AtCurrPrices("Стоимость материальных ресурсов (всего)")},
            OutputData{set_name: Some("РЕНЕЙМ................"), moving: Moving::Remain, expected_columns: 1, source: Source::AtCurrPrices("Производство работ в зимнее время 4%")},
            OutputData{set_name: Some("УДАЛИТЬ..............."), moving: Moving::Delete, expected_columns: 1, source: Source::AtCurrPrices("Производство работ в зимнее время 4%")},
        ];
    let printpart = PrintPart::new(vec_to_test);

    assert_eq!(3, printpart.get_number_of_columns());
}

pub struct Report<'a> {
    pub book: Option<xlsxwriter::Workbook>,
    pub part_1_just: PrintPart<'a>,
    pub part_2_base: PrintPart<'a>,
    pub part_3_curr: PrintPart<'a>,
}

impl<'a> Report<'a> {
    pub fn set_sample(sample: &Act) -> Result<Report, &'static str> {
        //-> Report
        // Нужно чтобы код назначал длину таблицы по горизонтали в зависимости от количества строк в итогах (обычно итоги имеют 17 строк,
        // но если какой-то акт имеет 16, 18, 0 или, скажем, 40 строк в итогах, то нужна какая-то логика, чтобы соотнести эти 40 строк одного акта
        // с 17 строками других актов. Нужно решение, как не сокращать эти 40 строк до 17 стандартных и выдать информацию пользователю без потерь.
        // Данные делятся на ожидаемые (им порядок можно сразу задать) и случайные.
        // Ниже массив, содержащий информацию о колонках, которые мы ожидаем получить из актов, здесь будем задавать порядок.
        // Позиция в массиве будет соответсвовать столбцу выходной формы (это крайние левые столбцы шапки):

        #[rustfmt::skip]
        let vec_1 = vec![
            OutputData{set_name: None,                                  moving: Moving::Remain, expected_columns: 1, source: Source::InTableHeader("Исполнитель")},
            OutputData{set_name: Some("Глава"),                         moving: Moving::Remain, expected_columns: 1, source: Source::Calculate},
            OutputData{set_name: None,                                  moving: Moving::Remain, expected_columns: 1, source: Source::InTableHeader("Объект")},
            OutputData{set_name: None,                                                        moving: Moving::Move, expected_columns: 1, source: Source::AtCurrPrices("Стоимость материальных ресурсов (всего)")},
            OutputData{set_name: None,                                  moving: Moving::Remain, expected_columns: 1, source: Source::InTableHeader("Договор №")},
            OutputData{set_name: None,                                  moving: Moving::Remain, expected_columns: 1, source: Source::InTableHeader("Договор дата")},
            OutputData{set_name: Some("Роботизация машин"),                                   moving: Moving::Move, expected_columns: 1, source: Source::AtBasePrices("Эксплуатация машин")},
            OutputData{set_name: None,                                  moving: Moving::Remain, expected_columns: 1, source: Source::InTableHeader("Смета №")},
            OutputData{set_name: None,                                  moving: Moving::Remain, expected_columns: 1, source: Source::InTableHeader("Смета наименование")},
            OutputData{set_name: Some("По смете в ц.2000г."),           moving: Moving::Remain, expected_columns: 1, source: Source::Calculate},
            OutputData{set_name: Some("Выполнение работ в ц.2000г."),   moving: Moving::Remain, expected_columns: 1, source: Source::Calculate},
            OutputData{set_name: None,                                  moving: Moving::Remain, expected_columns: 1, source: Source::InTableHeader("Акт №")},
            OutputData{set_name: Some("Акт дата"),                      moving: Moving::Remain, expected_columns: 1, source: Source::Calculate},
            OutputData{set_name: Some("Отчетный период начало"),        moving: Moving::Remain, expected_columns: 1, source: Source::Calculate},
            OutputData{set_name: Some("Отчетный период окончание"),     moving: Moving::Remain, expected_columns: 1, source: Source::Calculate},
            OutputData{set_name: None,                                  moving: Moving::Remain, expected_columns: 1, source: Source::InTableHeader("Метод расчета")},
            OutputData{set_name: Some("Ссылка на папку"),               moving: Moving::Remain, expected_columns: 1, source: Source::Calculate},
            OutputData{set_name: Some("Ссылка на файл"),                moving: Moving::Remain, expected_columns: 1, source: Source::Calculate},
            OutputData{set_name: Some("РЕНЕЙМ................"),                               moving: Moving::Remain, expected_columns: 1, source: Source::AtCurrPrices("Производство работ в зимнее время 4%")},
            OutputData{set_name: Some("УДАЛИТЬ..............."),                               moving: Moving::Delete, expected_columns: 1, source: Source::AtCurrPrices("Производство работ в зимнее время 4%")},
        ];
        // В векторе выше, перечислены далеко не все столбцы, что будут в акте (в акте может быть что угодно и при этом повторяться в неизвестном количестве).
        // В PART_1 мы перечислили, то чему хотели задать порядок заранее, но есть столбцы, где мы хотим оставить порядок, который существует в актах.
        // Чтобы продолжить, поделим отсутсвующие столбцы на два вида: соответсвующие форме акта, заданного в качестве шаблона, и те, которые в его форму не вписались.
        // Столбцы, которые будут совпадать со структурой шаблонного акта, получат приоритет и будут стремится в левое положение таблицы в том же порядке.
        // Другими словами, структура нашего отчета воспроизведет порядок итогов из шаблонного акта. Все что не вписальось в эту структуру будет размещено в крайних правых столбцах Excel.
        // У нас два вида данных в итогах: базовые и текущие цены, получается отчет будет написан из 3 частей.

        let part_1_just = PrintPart::new(vec_1.clone());
        let part_2_base = PrintPart::new(vec_1.clone());
        let part_3_curr = PrintPart::new(vec_1);

        Ok(Report {
            book: None,
            part_1_just,
            part_2_base,
            part_3_curr,
        })
    }
    fn other_parts<'b>(sample: &'b Act, part_1: &[OutputData<'a>]) -> Vec<OutputData<'a>> {
        let exclude_from_base = part_1
            .iter()
            .filter(|outputdata| match outputdata.source {
                Source::AtBasePrices(_) => true,
                _ => false,
            })
            .collect::<Vec<_>>();

        let exclude_from_curr = part_1
            .iter()
            .filter(|outputdata| match outputdata.source {
                Source::AtCurrPrices(_) => true,
                _ => false,
            })
            .collect::<Vec<_>>();

        let (part_2_base, part_3_curr) = sample.data_of_totals.iter().fold(
            (Vec::<OutputData>::new(), Vec::<OutputData>::new()),
            |mut acc, smpl_totalsrow| {
                let (check_renaming, not_listed, set_name) = exclude_from_base.iter().fold(
                    (false, true, None),
                    |(mut it_remains, mut not_listed, mut new_name), item| {
                        match item {
                            OutputData {
                                set_name,
                                moving: Moving::Remain,
                                source: Source::AtBasePrices(name),
                                ..
                            } if *name == smpl_totalsrow.name => {
                                it_remains = true;
                                not_listed = false;
                                new_name = Some(*name)
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

                if check_renaming || not_listed {
                    let columns_min = smpl_totalsrow.base_price.iter().map(Option::is_some).count();

                    let outputdata = OutputData {
                        set_name,
                        moving: Moving::Remain,
                        expected_columns: columns_min,
                        source: Source::AtBasePrices(&smpl_totalsrow.name),
                    };
                    acc.0.push(outputdata)
                }

                acc
            },
        );
        part_2_base
    }

    pub fn _write_as_sample(_book: xlsxwriter::Workbook) {}
    pub fn _format() {}
}
