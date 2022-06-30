#[derive(Debug)]
// pub enum ErrKind {
//     Handled,
//     Fatal,
// }

pub enum ErrName {
    ShiftedColumnsInHeader,
    SheetNotContainAllNecessaryData,
    CalamineSheetOfTheBookIsUndetectable,
    CalamineSheetOfTheBookIsUnreadable(calamine::XlsxError),
    Calamine,
}

#[derive(Debug)]
pub struct ErrDescription {
    pub name: ErrName,
    pub description: Option<String>,
}

pub fn error_message(err: ErrDescription, sh_name: &str) -> Option<String> {
    match err.name {
        ErrName::CalamineSheetOfTheBookIsUnreadable(_) => Some(format!("Какая-то проблема с чтением листа {}", sh_name)),
        ErrName::ShiftedColumnsInHeader => Some(String::from("Обнаружен нестандартный заголовок в Акте КС-2.\
        \nОжидаемая диспозиция столбцов для успешного сбора такова:
        \"Стройка\" и \"Объект\"                      - находятся в 1 столбце,
        \"Наименование работ и затрат\"             - находится в 4 столбце,
        \"Номер документа\"                         - находится в 6 столбце,
        \"Договор подряда\" и \"Доп. соглашение\"     - находятся в 10 столбце.")),
        ErrName::SheetNotContainAllNecessaryData => Some(String::from("В акте не полные данные.\
        \nОт файла требуется набор следующих ключевых слов:\
        \n  \"Стройка\",\
        \n  \"Объект\",\
        \n  \"Договор подряда\",\
        \n  \"Доп. соглашение\",\
        \n  \"Номер документа\",\
        \n  \"Наименование работ и затрат\",\
        \n  \"Стоимость материальных ресурсов (всего)\".\
        \n\nЕсли чего-то из перечисленного в акте не обнаружено, такой акт не может быть собран.\
        \nПроверьте документ на наличие перечисленных ключевых слов.\
        \nЕсли ошибка происходит при наличии всех ключевых слов - проверьте строковый порядок. \
        \nВхождение слов по строкам должно быть в порядке перечисления здесь: т.е. в файле\
        \nстрока \"Стройка\" должна быть выше строки \"Объект\", а \"Объект\", в свою очередь,\
        \nрасположен выше строки с текстом \"Договор подряда\".")),
        ErrName::CalamineSheetOfTheBookIsUndetectable => Some(format!("Встретился файл, который не содержит запрашиваемого вами листа \"{x}\".\n\
        Чтобы успешно выполнить процедуру сбора файлов, выполните одно из перечисленных действий над файлом, вызывающим ошибку:\n\
        - откройте файл и присвойте листу с актом имя \"{x}\";\n\
        - если вы не хотите собирать этот файл, переименуйте его, добавив к старому имени символ \"@\", или удалите файл из папки.", x = sh_name)),
        ErrName::Calamine => None,
    }
}
pub fn variant_eq<T>(first: &T, second: &T) -> bool {
    std::mem::discriminant(first) == std::mem::discriminant(second)
}
