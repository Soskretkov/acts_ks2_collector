use std::fmt;
use std::path::PathBuf;

#[derive(Debug)]
pub enum Error<'a> {
    InternalLogic(&'static str),
    NoFilesInSpecifiedPath(PathBuf),
    CalamineSheetOfTheBookIsUndetectable {
        sh_name_for_search: &'a str,
        sh_names: Vec<String>,
    },
    CalamineSheetOfTheBookIsUnreadable {
        sh_name: String, // нельзя ссылкой - имя листа с учетом регистра определяется внутри функции, где возможна ошибка
        err: calamine::XlsxError,
    },
    EmptySheetRange(&'a str),
    ShiftedColumnsInHeader,
    SheetNotContainAllNecessaryData,
    XlsxwriterWorkbookCreationError {
        wb_name: &'a str,
        err: xlsxwriter::XlsxError,
    },
    XlsxwriterSheetCreationFailed,
    XlsxwriterCellWriteFailed(xlsxwriter::XlsxError),
}

impl fmt::Display for Error<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InternalLogic(err) => {
                let msg = format!(
                    r#" Возникла ошибка во внутренней логики программы.
                
     Подробности об ошибке:
     {err}"#
                );
                write!(f, "{}", msg)
            }
            Self::NoFilesInSpecifiedPath(path) => {
                let path = path.display();
                let msg = format!(" Нет файлов \".xlsm\" по указанному пути: {path}");
                write!(f, "{}", msg)
            }
            Self::CalamineSheetOfTheBookIsUndetectable {
                sh_name_for_search,
                sh_names,
            } => {
                let string_sh_names = format!("{:?}", sh_names);

                // базовое сообщение
                let base_msg = format!(
                    r#" Встретился файл, который не содержит запрашиваемого вами листа "{sh_name_for_search}",
     так как файл имеет только следующие листы:
    
     {string_sh_names}
     
     "#
                );

                // дополнительное сообщение про кавычки как возможную причину ошибки
                let quotes_msg = if sh_name_for_search.starts_with('"')
                    && sh_name_for_search.ends_with('"')
                {
                    format!(
                        r#" Обратите внимание: вы ввели имя листа, заключённое в кавычки ("{sh_name_for_search}"), эти кавычки,
     могут являться причиной ошибки, так как обычно имена листов в книгах Excel не заключают в кавычки.
     Попробуйте повторить процедуру и ввести имя листа таким, каким вы его видите в самом файле Excel.
     
     "#
                    )
                } else {
                    "".to_string()
                };

                // заключительная часть сообщения
                let footer_msg = r#" Чтобы успешно выполнить процедуру сбора файлов, выполните одно из перечисленных действий:
     - откройте файл, вызывающий ошибку, и присвойте листу с актом имя, которое затем укажете программе;
     - если не хотите собирать этот файл, переименуйте файл, добавив к существующему имени символ "@",
       или удалите файл из папки;
     - если не хотите собирать папку, где находится файл, добавьте к существующему имени папки символ "@""#;

                // объединение всех частей в одно сообщение
                let full_msg = format!("{base_msg}{quotes_msg}{footer_msg}");
                write!(f, "{}", full_msg)
            }
            Self::CalamineSheetOfTheBookIsUnreadable { sh_name, err } => {
                let msg = format!(
                    r#" Возникла проблема с чтением листа {sh_name}.
                
     Подробности об ошибке:
     {err}"#
                );
                write!(f, "{}", msg)
            }
            Self::EmptySheetRange(sh_name) => {
                let msg = format!(" Лист {sh_name} не содержит данных (пуст)");
                write!(f, "{}", msg)
            }
            Self::ShiftedColumnsInHeader => write!(
                f,
                r#" Обнаружен нестандартный заголовок в Акте «КС-2».
     Ожидаемая диспозиция столбцов для успешного сбора такова:
            "Стройка" и "Объект"                      - находятся в одном столбце,
            "Наименование работ и затрат"             - смещение на 3 столбца  относительно "Стройки" и "Объекта",
            "Номер документа"                         - смещение на 5 столбцов относительно "Стройки" и "Объекта",
            "Договор подряда" и "Доп. соглашение"     - смещение на 9 столбцов относительно "Стройки" и "Объекта"."#,
            ),
            Self::SheetNotContainAllNecessaryData => write!(
                f,
                r#" В акте не полные данные.
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
     Если ошибка происходит при наличии всех ключевых слов - проверьте строковый порядок: 
     вхождение слов по строкам должно быть в порядке перечисленом выше (т.е. в файле
     строка "Стройка" должна быть выше строки с "Объект", а "Объект", в свою очередь,
     расположен выше строки с текстом "Договор подряда" и так далее)."#,
            ),
            Self::XlsxwriterWorkbookCreationError { wb_name, err } => {
                let msg = format!(
                    r#" Возникла проблема с созданием файла Excel с именем {wb_name}, который содержит результат
     работы программы.
                
     Подробности об ошибке:
     {err}"#
                );
                write!(f, "{}", msg)
            }
            Self::XlsxwriterSheetCreationFailed => {
                let msg = format!(
                    " Возникла проблема с созданием листа результата внутри нового файла Excel,
     на котором должен был размещен результат работы программы."
                );
                write!(f, "{}", msg)
            }
            Self::XlsxwriterCellWriteFailed(err) => {
                let msg = format!(
                    r#" Возникла проблема при попытке записи данных в ячейку нового файла Excel,
     того самого, который ожидается как результат работы программы.
                
     Подробности об ошибке:
     {err}"#
                );
                write!(f, "{}", msg)
            }
        }
    }
}
