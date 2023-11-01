use crate::config::XL_FILE_EXTENSION;
use std::fmt;
use std::path::PathBuf;

#[derive(Debug)]
pub enum Error<'a> {
    // InternalLogic(&'static str),
    NoFilesInSpecifiedPath(PathBuf),
    CalamineSheetOfTheBookIsUndetectable {
        file_path: &'a PathBuf,
        sh_name_for_search: &'a str,
        sh_names: Vec<String>,
    },
    CalamineSheetOfTheBookIsUnreadable {
        file_path: &'a PathBuf,
        sh_name: String, // нельзя ссылкой - имя листа с учетом регистра определяется внутри функции, где возможна ошибка
        err: calamine::XlsxError,
    },
    EmptySheetRange {
        file_path: &'a PathBuf,
        sh_name: String, // нельзя ссылкой - имя листа с учетом регистра определяется внутри функции, где возможна ошибка
    },
    ShiftedColumnsInHeader(&'a PathBuf),
    SheetNotContainAllNecessaryData(&'a PathBuf),
    XlsxwriterWorkbookCreation {
        wb_name: &'a str,
        err: xlsxwriter::XlsxError,
    },
    XlsxwriterSheetCreationFailed,
    XlsxwriterCellWriteFailed(xlsxwriter::XlsxError),
    XlsxwriterWorkbookClose {
        wb_name: &'a str,
        err: xlsxwriter::XlsxError,
    },
}

impl<'a> std::error::Error for Error<'a> {}

impl fmt::Display for Error<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            //         Self::InternalLogic(err) => {
            //             let msg = format!(
            //                 r#" Во внутренней логике программы произошла ошибка.

            //  Подробности об ошибке:
            //  {err}"#
            //             );
            //             write!(f, "{msg}")
            //         }
            Self::NoFilesInSpecifiedPath(path) => {
                let path = path.display();
                let msg = format!(" Нет файлов \"{XL_FILE_EXTENSION}\" по указанному пути: {path}");
                write!(f, "{msg}")
            }
            Self::CalamineSheetOfTheBookIsUndetectable {
                file_path,
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

                // опциональное сообщение про кавычки как возможную причину ошибки
                let optional_msg = if sh_name_for_search.starts_with('"')
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
 - если не хотите собирать папку, где находится файл, добавьте к существующему имени папки символ "@"
"#;

                let path_msg = format!(" Файл, вызывающий ошибку:\n{}", file_path.display());

                // объединение всех частей в одно сообщение
                let full_msg = format!("{base_msg}\n{optional_msg}\n{footer_msg}\n{path_msg}");
                write!(f, "{full_msg}")
            }
            Self::CalamineSheetOfTheBookIsUnreadable {
                file_path,
                sh_name,
                err,
            } => {
                let base_msg = format!(" Возникла проблема с чтением листа «{sh_name}».\n");
                let footer_msg = format!(" Подробности об ошибке:\n{err}\n");
                let path_msg = format!(" Файл, вызывающий ошибку:\n{}", file_path.display());
                let full_msg = format!("{base_msg}\n{footer_msg}\n{path_msg}");
                write!(f, "{full_msg}")
            }
            Self::EmptySheetRange { file_path, sh_name } => {
                let base_msg = format!(" Лист «{sh_name}» не содержит данных (пуст)\n");
                let path_msg = format!(" Файл, вызывающий ошибку:\n{}", file_path.display());
                let full_msg = format!("{base_msg}\n{path_msg}");
                write!(f, "{full_msg}")
            }
            Self::ShiftedColumnsInHeader(file_path) => {
                let base_msg = r#" Обнаружен нестандартный заголовок в Акте «КС-2».
 Ожидаемая диспозиция столбцов для успешного сбора такова:
    "Стройка" и "Объект"                      - находятся в одном столбце,
    "Наименование работ и затрат"             - смещение на 3 столбца  относительно "Стройки" и "Объекта",
    "Номер документа"                         - смещение на 5 столбцов относительно "Стройки" и "Объекта",
    "Договор подряда" и "Доп. соглашение"     - смещение на 9 столбцов относительно "Стройки" и "Объекта".
"#;
                let path_msg = format!(" Файл, вызывающий ошибку:\n{}", file_path.display());
                let full_msg = format!("{base_msg}\n{path_msg}");
                write!(f, "{full_msg}")
            }
            Self::SheetNotContainAllNecessaryData(file_path) => {
                let base_msg = r#" В акте не полные данные.
 От собираемого файла требуется следующий набор ключевых слов:
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
 расположен выше строки с текстом "Договор подряда" и так далее).
"#;
                let path_msg = format!(" Файл, вызывающий ошибку:\n{}", file_path.display());
                let full_msg = format!("{base_msg}\n{path_msg}");
                write!(f, "{full_msg}")
            }
            Self::XlsxwriterWorkbookCreation { wb_name, err } => {
                let msg = format!(
                    r#" Не удалась попытка создания файла Excel с именем {wb_name}, речь о файле Excel,
 который содержит результат работы программы.
                
 Подробности об ошибке:
 {err}"#
                );
                write!(f, "{msg}")
            }
            Self::XlsxwriterSheetCreationFailed => {
                let msg = format!(
                    " Не удалась попытка создание листа результата внутри нового файла Excel, речь о листе Excel на котором
 должен был быть записан результат работы программы."
                );
                write!(f, "{msg}")
            }
            Self::XlsxwriterCellWriteFailed(err) => {
                let msg = format!(
                    r#" Не удалась попытка записи данных в ячейку нового файла Excel, того самого,
 который ожидается как результат работы программы.
                
     Подробности об ошибке:
     {err}"#
                );
                write!(f, "{msg}")
            }
            Self::XlsxwriterWorkbookClose { wb_name, .. } => {
                let msg = format!(
                    r#" Не удалось сохранение на диск файла Excel с именем {wb_name}, который содержит
 результат работы программы.
 
 Вероятная причина: не закрыт файл Excel с результатами прошлого сбора."#
                );
                write!(f, "{msg}")
            }
        }
    }
}
