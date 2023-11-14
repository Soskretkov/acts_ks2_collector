use crate::constants::XL_FILE_EXTENSION;
use crate::errors::Error;
use crate::ui;
use calamine::{Xlsx, XlsxError};
use std::fs::File;
use std::io::BufReader;
use std::path::PathBuf;
use walkdir::WalkDir;
pub struct Book {
    pub path: PathBuf,
    pub data: Xlsx<BufReader<File>>,
}

impl Book {
    pub fn new(path: PathBuf) -> Result<Self, XlsxError> {
        let data: Xlsx<_> = calamine::open_workbook(&path)?;
        Ok(Book { path, data })
    }
}

pub struct ExtractedBooks {
    pub books: Vec<Result<Book, XlsxError>>,
    pub file_count_excluded: usize,
}

impl ExtractedBooks {
    pub fn new(path: &PathBuf) -> Result<Self, Error<'static>> {
        let files: Vec<_> = WalkDir::new(path)
            .into_iter()
            .filter_map(|e| e.ok()) //будет молча пропускать каталоги, на доступ к которым у владельца запущенного процесса нет разрешения
            .filter(|e| {
                e.file_name()
                    .to_str()
                    .map(|s| !s.starts_with('~') & s.ends_with(XL_FILE_EXTENSION))
                    .unwrap_or_else(|| false)
            })
            .collect();

        let mut xl_files_vec = vec![];
        let mut file_count_excluded = 0;
        let mut file_print_counter = 0;

        for entry in files {
            let file_checked_path = entry
                .path()
                .strip_prefix(path)
                .map_err(|err| Error::InternalLogic {
                    tech_descr: format!(
                        r#"Не удалось выполнить проверку на наличие символа "@" в пути для файла:
{}"#,
                        entry.path().to_string_lossy()
                    ),
                    err: Some(Box::new(err)),
                })?
                .to_string_lossy();

            if path.is_dir() {
                if file_checked_path.contains('@') {
                    file_count_excluded += 1;
                    continue;
                }
            }

            if xl_files_vec.len() == 0 {
                ui::display_formatted_text("\nОтбранны файлы:", None);
            }

            let file_display_path = if path.is_dir() {
                file_checked_path
            } else {
                path.to_string_lossy()
            };

            file_print_counter += 1;
            let msg = format!("{}: {}", file_print_counter, file_display_path);
            ui::display_formatted_text(&msg, None);

            let xl_file = Book::new(entry.into_path());
            xl_files_vec.push(xl_file);
        }

        Ok(Self {
            books: xl_files_vec,
            file_count_excluded,
        })
    }
}
