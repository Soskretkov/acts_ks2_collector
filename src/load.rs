use crate::transform::Act;

#[derive(Debug, Clone)]
pub struct OutputData {
    pub new_name: Option<&'static str>,
    pub number_of_copies: usize,
    pub data_source: DataSource,
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
            .fold(0, |acc, copy| acc + copy.number_of_copies)
    }
}
#[test]
fn PrintPart_test() {
    #[rustfmt::skip]
        let vec_to_test = vec![
            OutputData{new_name: None,           number_of_copies: 1,  data_source: DataSource::InTableHeader("Исполнитель")},
            OutputData{new_name: Some("Глава"),  number_of_copies: 11, data_source: DataSource::Calculate},
            OutputData{new_name: None,           number_of_copies: 0,  data_source: DataSource::InTableHeader("Объект")},
        ];
    let printpart = PrintPart::new(vec_to_test);

    assert_eq!(12, printpart.get_number_of_columns());
}

pub struct Report {
    pub book: Option<xlsxwriter::Workbook>,
    pub part_1: PrintPart,
    pub part_2: PrintPart,
    pub part_3: PrintPart,
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
            OutputData{new_name: None,                                  number_of_copies: 1, data_source: DataSource::InTableHeader("Исполнитель")},
            OutputData{new_name: Some("Глава"),                         number_of_copies: 1, data_source: DataSource::Calculate},
            OutputData{new_name: None,                                  number_of_copies: 1, data_source: DataSource::InTableHeader("Объект")},
            OutputData{new_name: None,                                                        number_of_copies: 1, data_source: DataSource::AtCurrPrices("Стоимость материальных ресурсов (всего)")},
            OutputData{new_name: None,                                  number_of_copies: 1, data_source: DataSource::InTableHeader("Договор №")},
            OutputData{new_name: None,                                  number_of_copies: 1, data_source: DataSource::InTableHeader("Договор дата")},
            OutputData{new_name: Some("Роботизация машин"),                                   number_of_copies: 1, data_source: DataSource::AtBasePrices("Эксплуатация машин")},
            OutputData{new_name: None,                                  number_of_copies: 1, data_source: DataSource::InTableHeader("Смета №")},
            OutputData{new_name: None,                                  number_of_copies: 1, data_source: DataSource::InTableHeader("Смета наименование")},
            OutputData{new_name: Some("По смете в ц.2000г."),           number_of_copies: 1, data_source: DataSource::Calculate},
            OutputData{new_name: Some("Выполнение работ в ц.2000г."),   number_of_copies: 1, data_source: DataSource::Calculate},
            OutputData{new_name: None,                                  number_of_copies: 1, data_source: DataSource::InTableHeader("Акт №")},
            OutputData{new_name: Some("Акт дата"),                      number_of_copies: 1, data_source: DataSource::Calculate},
            OutputData{new_name: Some("Отчетный период начало"),        number_of_copies: 1, data_source: DataSource::Calculate},
            OutputData{new_name: Some("Отчетный период окончание"),     number_of_copies: 1, data_source: DataSource::Calculate},
            OutputData{new_name: None,                                  number_of_copies: 1, data_source: DataSource::InTableHeader("Метод расчета")},
            OutputData{new_name: Some("Ссылка на папку"),               number_of_copies: 1, data_source: DataSource::Calculate},
            OutputData{new_name: Some("Ссылка на файл"),                number_of_copies: 1, data_source: DataSource::Calculate},
        ];
        // В векторе REPORTING_PRESETS выше, перечислены далеко не все столбцы, что будут в акте (в акте может быть что угодно и при этом повторяться в неизвестном количестве).
        // В PART_1 войдет то, что мы перечислили, чему хотели задать порядок заранее, но есть столбцы, где мы хотим оставить порядок, который существует в актах.
        // Поделим отсутсвующие столбцы на два вида: соответсвующие форме акта, заданного в качестве шаблона, и те, которые в его форму не вписались.
        // Столбцы, которые будут совпадать со структурой шаблонного акта, получат приоритет и будут стремится в левое положение таблицы в том же порядке.
        // Другими словами, структура нашего отчета воспроизведет порядок итогов из шаблонного акта. Все что не вписальось в эту структуру будет размещено в крайних правых столбцах Excel.
        // У нас два вида данных в итогах: базовые и текущие цены, получается отчет будет написан из 3 частей.

        let part_1 = PrintPart::new(vec_1.clone());
        let part_2 = PrintPart::new(vec_1.clone());
        let part_3 = PrintPart::new(vec_1);

        Ok(Report {
            book: None,
            part_1,
            part_2,
            part_3,
        })
    }
    pub fn _write_as_sample(_book: xlsxwriter::Workbook) {}
    pub fn _format() {}
}

// pub fn first_file_data_names(act: &[TotalsRow]) -> (Vec<&String>, Vec<&String>) {
//     let (already_collected_base, already_collected_curr) =
//         PART_1_REPORT
//             .iter()
//             .fold((Vec::new(), Vec::new()), |mut acc, (_, source)| {
//                 if let DataSource::AtBasePrices(default_name) = source {
//                     acc.0.push(*default_name)
//                 };

//                 if let DataSource::AtCurrPrices(default_name) = source {
//                     acc.1.push(*default_name)
//                 };
//                 acc
//             });

//     let (fst_fls_base, fst_fls_curr) =
//         act.iter()
//             .fold((Vec::new(), Vec::new()), |mut acc, x| {
//                 if x.base_price.is_some() && !already_collected_base.iter().any(|item| item == &x.name) {
//                         acc.0.push(&x.name)
//                 }

//                 if x.current_price.is_some() && !already_collected_curr.iter().any(|item| item == &x.name) {
//                         acc.1.push(&x.name)

//                 }

//                 acc
//             });

//     (fst_fls_base, fst_fls_curr)
// }
