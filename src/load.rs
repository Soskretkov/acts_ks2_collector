use crate::transform::Act;

#[derive(Debug, Clone)]
pub struct OutputData<'a> {
    pub set_name: Option<&'a str>,
    // pub change_location: Parts,
    pub print_quantity: Option<usize>,
    pub source: DataSource<'a>,
}
// Четыре вида данных на выходе: в готовом виде в шапке, в готов виде в итогах акта (2 варанта), и нет готовых (нужно расчитать программой):
#[derive(Debug, Clone, PartialEq)]

enum Parts {
    First(usize),
    Curr,
    Base,
}

#[derive(Debug, Clone)]
pub enum DataSource<'a> {
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
        vector.iter().fold(0, |acc, copy| match copy.print_quantity {
            Some(x) => acc + x,
            None => acc,
        })
    }
}
#[test]
fn PrintPart_test() {
    #[rustfmt::skip]
        let vec_to_test = vec![
            OutputData{set_name: None,           print_quantity: Some(1),  source: DataSource::InTableHeader("Исполнитель")},
            OutputData{set_name: Some("Глава"),  print_quantity: Some(11), source: DataSource::Calculate},
            OutputData{set_name: None,           print_quantity: Some(0),  source: DataSource::InTableHeader("Объект")},
        ];
    let printpart = PrintPart::new(vec_to_test);

    assert_eq!(12, printpart.get_number_of_columns());
}

pub struct Report<'a> {
    pub book: Option<xlsxwriter::Workbook>,
    pub part_1_just: PrintPart<'a>,
    pub part_2_base: PrintPart<'a>,
    pub part_3_curr: PrintPart<'a>,
}

impl<'a> Report<'a> {
    pub fn set_sample(sample: &Act) -> Result<Report, &'static str> {
        //-> Report{
        // Нужно чтобы код назначал длину таблицы по горизонтали в зависимости от количества строк в итогах (обычно итоги имеют 17 строк,
        // но если какой-то акт имеет 16, 18, 0 или, скажем, 40 строк в итогах, то нужна какая-то логика, чтобы соотнести эти 40 строк одного акта
        // с 17 строками других актов. Нужно решение, как не сокращать эти 40 строк до 17 стандартных и выдать информацию пользователю без потерь.
        // Данные делятся на ожидаемые (им порядок можно сразу задать) и случайные.
        // Ниже массив, содержащий информацию о колонках, которые мы ожидаем получить из актов, здесь будем задавать порядок.
        // Позиция в массиве будет соответсвовать столбцу выходной формы (это крайние левые столбцы шапки):

        #[rustfmt::skip]
        let vec_1 = vec![
            OutputData{set_name: None,                                  print_quantity: Some(1), source: DataSource::InTableHeader("Исполнитель")},
            OutputData{set_name: Some("Глава"),                         print_quantity: Some(1), source: DataSource::Calculate},
            OutputData{set_name: None,                                  print_quantity: Some(1), source: DataSource::InTableHeader("Объект")},
            OutputData{set_name: None,                                                        print_quantity: Some(1), source: DataSource::AtCurrPrices("Стоимость материальных ресурсов (всего)")},
            OutputData{set_name: None,                                  print_quantity: Some(1), source: DataSource::InTableHeader("Договор №")},
            OutputData{set_name: None,                                  print_quantity: Some(1), source: DataSource::InTableHeader("Договор дата")},
            OutputData{set_name: Some("Роботизация машин"),                                   print_quantity: Some(1), source: DataSource::AtBasePrices("Эксплуатация машин")},
            OutputData{set_name: None,                                  print_quantity: Some(1), source: DataSource::InTableHeader("Смета №")},
            OutputData{set_name: None,                                  print_quantity: Some(1), source: DataSource::InTableHeader("Смета наименование")},
            OutputData{set_name: Some("По смете в ц.2000г."),           print_quantity: Some(1), source: DataSource::Calculate},
            OutputData{set_name: Some("Выполнение работ в ц.2000г."),   print_quantity: Some(1), source: DataSource::Calculate},
            OutputData{set_name: None,                                  print_quantity: Some(1), source: DataSource::InTableHeader("Акт №")},
            OutputData{set_name: Some("Акт дата"),                      print_quantity: Some(1), source: DataSource::Calculate},
            OutputData{set_name: Some("Отчетный период начало"),        print_quantity: Some(1), source: DataSource::Calculate},
            OutputData{set_name: Some("Отчетный период окончание"),     print_quantity: Some(1), source: DataSource::Calculate},
            OutputData{set_name: None,                                  print_quantity: Some(1), source: DataSource::InTableHeader("Метод расчета")},
            OutputData{set_name: Some("Ссылка на папку"),               print_quantity: Some(1), source: DataSource::Calculate},
            OutputData{set_name: Some("Ссылка на файл"),                print_quantity: Some(1), source: DataSource::Calculate},
            OutputData{set_name: None,                                                        print_quantity: None,    source: DataSource::AtCurrPrices("Итого с К = 1")},
            OutputData{set_name: Some("Итого с К = 9999999999999999"),                        print_quantity: Some(1), source: DataSource::AtBasePrices("Итого с К = 1")},
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
    fn other_parts<'b>(sample: &'b Act, part_1: &[OutputData]) {
        let (exclude_from_base, exclude_from_curr_part) =
            part_1
                .iter()
                .fold((Vec::new(), Vec::new()), |mut acc, outputdata| {
                    if let DataSource::AtBasePrices(default_name) = outputdata.source {
                        acc.0.push(default_name)
                    };

                    if let DataSource::AtCurrPrices(default_name) = outputdata.source {
                        acc.1.push(default_name)
                    };
                    acc
                });

        let (part_2_base, part_3_curr) = sample.data_of_totals.iter().fold(
            (Vec::<OutputData>::new(), Vec::<OutputData>::new()),
            |mut acc, x| {
                if !exclude_from_base.iter().any(|item| item == &x.name) {
                    let columns_min = x.base_price.iter().map(Option::is_some).count();

                    if columns_min > 0 {
                        let outputdata = OutputData {
                            set_name: None,
                            print_quantity: Some(columns_min),
                            source: DataSource::AtBasePrices(&x.name),
                        };
                        acc.0.push(outputdata)
                    }
                }

                acc
            },
        );
    }

    pub fn _write_as_sample(_book: xlsxwriter::Workbook) {}
    pub fn _format() {}
}
