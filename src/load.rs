use crate::transform::TotalsRow;

// Четыре вида данных на выходе: в готовом виде в шапке, в готов виде в итогах акта (2 варанта), и нет готовых и нужно расчитать программой:
#[derive(PartialEq)]
pub enum DataSource {
    InTableHeader(&'static str),
    AtCurrPrices(&'static str),
    AtBasePrices(&'static str),
    Calculate,
}

// Нужен код, который будет назначать длину таблицы по горизонтали в зависимости от количества строк в итогах (обычно итоги имеют 17 строк,
// но если какой-то акт имеет 16, 18, 0 или, скажем, 40 строк в итогах, то нужна какая-то логика, чтобы соотнести эти 40 строк одного акта
// с 17 строками других актов. Нужно решение, как не сокращать эти 40 строк до 17 стандартных и выдать информацию пользователю без потерь.
// Таким образом у нас данные условно делятся на ожидаемые (им порядок можно сразу задать) и случайные

// Нам нужна структура, содержащая информацию о колонках, которые мы ожидаем получить из актов, и здесь мы будем задавать порядок.
// Мы можем составить представление о начале заголовка выходной формы, читая кортеж схематично: ("Нужно ли переименовать имя по умолчанию?", "Где искать данные?"):
// Позиция кортежа в массиве будет соответсвовать столбцу выходной формы (это крайние левые столбцы шапки):

#[rustfmt::skip]
pub const PART_1_REPORT: [(Option<&'static str>, DataSource); 18] = [
    (None,                                  DataSource::InTableHeader("Исполнитель")),
    (Some("Глава"),                         DataSource::Calculate),
    (None,                                  DataSource::InTableHeader("Объект")),
    (None,                                                          DataSource::AtCurrPrices("Стоимость материальных ресурсов (всего)")),
    (None,                                  DataSource::InTableHeader("Договор №")),
    (None,                                  DataSource::InTableHeader("Договор дата")),
    (Some("Прихватизация машин"),                                   DataSource::AtBasePrices("Эксплуатация машин")),
    (None,                                  DataSource::InTableHeader("Смета №")),
    (None,                                  DataSource::InTableHeader("Смета наименование")),
    (Some("По смете в ц.2000г."),           DataSource::Calculate),
    (Some("Выполнение работ в ц.2000г."),   DataSource::Calculate),
    (None,                                  DataSource::InTableHeader("Акт №")),
    (Some("Акт дата"),                      DataSource::Calculate),
    (Some("Отчетный период начало"),        DataSource::Calculate),
    (Some("Отчетный период окончание"),     DataSource::Calculate),
    (None,                                  DataSource::InTableHeader("Метод расчета")),
    (Some("Ссылка на папку"),               DataSource::Calculate),
    (Some("Ссылка на файл"),                DataSource::Calculate),
];

// В массиве выше перечислены далеко не все столбцы что будут в акте (в акте может быть все что угодно и повторяться в неизвестном количестве).
// В PART_1 мы перечислили только то, чему хотели задать порядок заранее, но есть столбцы, где мы хотим оставить тот порядок, который существует в актах.
// Поделим отсутсвующие столбцы на два вида: соответсвующие форме акта первого в выборке и те, которые в его форму не вписались.
// Столбцы, которые будут совпадать со структурой первого акта, получат больший приоритет и будут стремится в левое положение таблицы.
// Другими словами, структура нашего отчета воспроизведет порядок итогов первого акта в выборке. А все что не вписальось в эту структуру будет помещено в крайние правые столбцы.
// У нас два вида данных в итогах: базовые и текущие цены. Это нужно учитывать.

pub fn first_file_data_names(act: &[TotalsRow]) -> (Vec<&String>, Vec<&String>) {
    let (already_collected_base, already_collected_curr) =
        PART_1_REPORT
            .iter()
            .fold((Vec::new(), Vec::new()), |mut acc, (_, source)| {
                if let DataSource::AtBasePrices(default_name) = source {
                    acc.0.push(*default_name)
                };

                if let DataSource::AtCurrPrices(default_name) = source {
                    acc.1.push(*default_name)
                };
                acc
            });

    let (fst_fls_base, fst_fls_curr) =
        act.iter()
            .fold((Vec::new(), Vec::new()), |mut acc, x| {
                if x.base_price.is_some() {
                    if x.duplicate_number.is_some() || !already_collected_base.iter().any(|item| item == &x.name) {
                        acc.0.push(&x.name)
                    }
                }
                
                if x.current_price.is_some() {
                    if x.duplicate_number.is_some() || !already_collected_curr.iter().any(|item| item == &x.name) {
                        acc.1.push(&x.name)                        
                    }
                }
                
                acc
            });

    (fst_fls_base, fst_fls_curr)
}
