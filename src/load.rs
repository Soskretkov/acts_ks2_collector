use crate::transform::Act;

#[derive(Debug, Clone)]
pub struct OutputData {
    pub rename: Option<&'static str>,
    pub columns_min: Option<usize>,
    pub source: DataSource,
}
// Четыре вида данных на выходе: в готовом виде в шапке, в готов виде в итогах акта (2 варанта), и нет готовых (нужно расчитать программой):
#[derive(Debug, Clone, PartialEq)]
pub enum DataSource {
    InTableHeader(&'static str),
    AtCurrPrices(&'static str),
    AtBasePrices(&'static str),
    Calculate,
}

pub struct PrintPart {
    vector: Vec<OutputData>,
    total_col: usize,
}

impl PrintPart {
    pub fn new(vector: Vec<OutputData>) -> PrintPart {
        let total_col = Self::count_col(&vector);

        PrintPart { vector, total_col }
    }
    pub fn get_number_of_columns(&self) -> usize {
        self.total_col
    }
    fn count_col(vector: &[OutputData]) -> usize {
        vector
            .iter()
            .fold(0, |acc, copy| match copy.columns_min {
                Some(x) => acc + x,
                None => acc,
            })
    }
}
#[test]
fn PrintPart_test() {
    #[rustfmt::skip]
        let vec_to_test = vec![
            OutputData{rename: None,           columns_min: Some(1),  source: DataSource::InTableHeader("Исполнитель")},
            OutputData{rename: Some("Глава"),  columns_min: Some(11), source: DataSource::Calculate},
            OutputData{rename: None,           columns_min: Some(0),  source: DataSource::InTableHeader("Объект")},
        ];
    let printpart = PrintPart::new(vec_to_test);

    assert_eq!(12, printpart.get_number_of_columns());
}

pub struct Report {
    pub book: Option<xlsxwriter::Workbook>,
    pub part_1_just: PrintPart,
    pub part_2_base: PrintPart,
    pub part_3_curr: PrintPart,
}

impl Report {
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
            OutputData{rename: None,                                  columns_min: Some(1), source: DataSource::InTableHeader("Исполнитель")},
            OutputData{rename: Some("Глава"),                         columns_min: Some(1), source: DataSource::Calculate},
            OutputData{rename: None,                                  columns_min: Some(1), source: DataSource::InTableHeader("Объект")},
            OutputData{rename: None,                                                        columns_min: Some(1), source: DataSource::AtCurrPrices("Стоимость материальных ресурсов (всего)")},
            OutputData{rename: None,                                  columns_min: Some(1), source: DataSource::InTableHeader("Договор №")},
            OutputData{rename: None,                                  columns_min: Some(1), source: DataSource::InTableHeader("Договор дата")},
            OutputData{rename: Some("Роботизация машин"),                                   columns_min: Some(1), source: DataSource::AtBasePrices("Эксплуатация машин")},
            OutputData{rename: None,                                  columns_min: Some(1), source: DataSource::InTableHeader("Смета №")},
            OutputData{rename: None,                                  columns_min: Some(1), source: DataSource::InTableHeader("Смета наименование")},
            OutputData{rename: Some("По смете в ц.2000г."),           columns_min: Some(1), source: DataSource::Calculate},
            OutputData{rename: Some("Выполнение работ в ц.2000г."),   columns_min: Some(1), source: DataSource::Calculate},
            OutputData{rename: None,                                  columns_min: Some(1), source: DataSource::InTableHeader("Акт №")},
            OutputData{rename: Some("Акт дата"),                      columns_min: Some(1), source: DataSource::Calculate},
            OutputData{rename: Some("Отчетный период начало"),        columns_min: Some(1), source: DataSource::Calculate},
            OutputData{rename: Some("Отчетный период окончание"),     columns_min: Some(1), source: DataSource::Calculate},
            OutputData{rename: None,                                  columns_min: Some(1), source: DataSource::InTableHeader("Метод расчета")},
            OutputData{rename: Some("Ссылка на папку"),               columns_min: Some(1), source: DataSource::Calculate},
            OutputData{rename: Some("Ссылка на файл"),                columns_min: Some(1), source: DataSource::Calculate},
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
    fn other_parts(sample: &Act, part_1: &[OutputData]) {
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
        // let (fst_fls_base, fst_fls_curr) = sample.data_of_totals.iter().fold(
        //     (Vec::<OutputData>::new(), Vec::<OutputData>::new()),
        //     |mut acc, x| {
        //         if !already_collected_base.iter().any(|item| item == &x.name) {
        //             let len_of_base = x.base_price.len();

        //             if len_of_base > 0 {
        //                 let outputdata = OutputData {
        //                     new_name: None,
        //                     number_of_copies: len_of_base,
        //                     data_source: DataSource::AtBasePrices(&x.name),
        //                 };
        //                 acc.0.push(outputdata)
        //             }
        //         }

        //         acc
        //     },
        // );
    }

    pub fn _write_as_sample(_book: xlsxwriter::Workbook) {}
    pub fn _format() {}
}