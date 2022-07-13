use std::path::PathBuf;
#[derive(Debug)]
pub enum ErrName {
    ShiftedColumnsInHeader,
    SheetNotContainAllNecessaryData,
    CalamineSheetOfTheBookIsUndetectable,
    CalamineSheetOfTheBookIsUnreadable(calamine::XlsxError),
    Calamine,
    NoFilesInSpecifiedPath(PathBuf),
}

#[derive(Debug)]
pub struct ErrDescription {
    pub name: ErrName,
    pub description: Option<String>,
}

pub fn error_message(err: ErrDescription, sh_name: &str) -> Option<String> {
    match err {
        ErrDescription {
            name: ErrName::CalamineSheetOfTheBookIsUnreadable(_),
            ..
        } => Some(format!("Какая-то проблема с чтением листа {}", sh_name)),
        ErrDescription {
            name: ErrName::ShiftedColumnsInHeader,
            ..
        } => Some(String::from(
            r#"Обнаружен нестандартный заголовок в Акте КС-2.
Ожидаемая диспозиция столбцов для успешного сбора такова:
        "Стройка" и "Объект"                      - находятся в 1 столбце,
        "Наименование работ и затрат"             - находится в 4 столбце,
        "Номер документа"                         - находится в 6 столбце,
        "Договор подряда" и "Доп. соглашение"     - находятся в 10 столбце."#,
        )),
        ErrDescription {
            name: ErrName::SheetNotContainAllNecessaryData,
            ..
        } => Some(String::from(
            r#"В акте не полные данные.
От файла требуется набор следующих ключевых слов:
  "Стройка",
  "Объект",
  "Договор подряда",
  "Доп. соглашение",
  "Номер документа",
  "Наименование работ и затрат",
  "Стоимость материальных ресурсов (всего)".

Если чего-то из перечисленного в акте не обнаружено, такой акт не может быть собран.
Проверьте документ на наличие перечисленных ключевых слов.
Если ошибка происходит при наличии всех ключевых слов - проверьте строковый порядок. 
Вхождение слов по строкам должно быть в порядке перечисленом выше (т.е. в файле
строка "Стройка" должна быть выше строки "Объект", а "Объект", в свою очередь,
расположен выше строки с текстом "Договор подряда")."#,
        )),
        ErrDescription {
            name: ErrName::CalamineSheetOfTheBookIsUndetectable,
            description: descr,
        } => Some(
            format!(
                r#"Встретился файл, который не содержит запрашиваемого вами листа "{x}",
так как файл имеет только следующие листы:

{y}"#,
                x = sh_name,
                y = descr.unwrap()
            ) + &if sh_name.starts_with('"') && sh_name.ends_with('"') {
                format!(
                    r#"

Обратите внимание, что введенное вами имя листа, {}, содержит кавычки;
эти кавычки, вероятно, являются причиной ошибки, вам следует ввести имя листа так,
как вы бы его обнаружили открыв файл Excel, не заключая текст ввода в кавычки.
"#,
                    sh_name
                )
            } else {
                "".to_string()
            } + r#"
Чтобы успешно выполнить процедуру сбора файлов, выполните одно из перечисленных действий:
- откройте файл, вызывающий ошибку, и присвойте листу с актом имя, которое затем укажете программе;
- если не хотите собирать этот файл, переименуйте файл, добавив к существующему имени символ "@", или удалите файл из папки;
- если не хотите собирать папку, где находится файл, добавьте к существующему имени папки символ "@""#,
        ),
        ErrDescription {
            name: ErrName::Calamine,
            ..
        } => None,

        ErrDescription {
            name: ErrName::NoFilesInSpecifiedPath(path),
            ..
        } => Some(format!(
            "Нет файлов \".xlsm\" по указанному пути: {}",
            path.display()
        )),
    }
}
pub fn variant_eq<T>(first: &T, second: &T) -> bool {
    std::mem::discriminant(first) == std::mem::discriminant(second)
}
