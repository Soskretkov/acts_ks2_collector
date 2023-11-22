use console::{Style, Term}; // для очистки консоли перед выводом полезных сообщений
use std::env;
use std::path;
use std::thread; // для засыпания на секунду-две при печати сообщений
use std::time::Duration; // для засыпания на секунду-две при печати сообщений // имя ".exe" будет присвоено файлу Excel
mod shared;
mod errors;
mod extract;
mod load;
mod ui;
use crate::shared::constants::{SUCCESS_PAUSE_DURATION, XL_FILE_EXTENSION};
use crate::errors::Error;
use crate::extract::Act;
use crate::extract::{ExtractedBooks, Sheet};
use crate::load::Report;

fn main() {
    Term::stdout().set_title("«Ks2 etl»,  v".to_string() + env!("CARGO_PKG_VERSION"));
    ui::display_first_lines(true);
    ui::display_help();
    let cyan = Style::new().cyan();
    let red = Style::new().red();
    'main_loop: loop {
        let (path, user_entered_sh_name) = match ui::user_input() {
            Ok(x) => x,
            Err(err) => {
                display_error_and_wait(err);
                continue 'main_loop;
            }
        };
        // Debug
        // let user_entered_sh_name = "Лист1".to_owned();
        // let path = std::path::PathBuf::from(r"C:\Users\User\rust\ks2_etl\02-01.1-0239С-2С-И3-17-01-2023 - копия.xlsm".to_string());
        // let path = std::path::PathBuf::from(r"C:\Users\User\rust\ks2_etl\02-01.1-0239С-2С-И3-17-01-2023 — копия.xlsm".to_string());

        let string_report_path = env::args()
            .next()
            .unwrap()
            .trim_end_matches(".exe")
            .to_owned()
            + ".xlsx";

        let report_path = path::PathBuf::from(string_report_path);

        let wraped_books_vec = ExtractedBooks::new(&path).map(|extracted_xl_books| {
            if path.is_dir() {
                let file_count_total =
                    extracted_xl_books.books.len() + extracted_xl_books.file_count_excluded;
                let base_msg = format!(
                    "Обнаружено {} файлов с расширением \"{}\".",
                    file_count_total, XL_FILE_EXTENSION
                );

                let footer_msg = if extracted_xl_books.file_count_excluded > 0 {
                    format!(
                        r#"Из них {} помечены "@" для исключения."#,
                        extracted_xl_books.file_count_excluded
                    )
                } else {
                    "Среди них нет файлов, помеченных как исключенные.".to_string()
                };

                let full_msg = format!(
                    "\n{}\n{}",
                    base_msg,
                    if file_count_total == 0 {
                        "".to_string()
                    } else {
                        footer_msg
                    }
                );

                ui::display_formatted_text(&full_msg, None);
            }

            extracted_xl_books.books
        });

        let books_vec = match wraped_books_vec {
            Ok(books_vec) => books_vec,
            Err(err) => {
                display_error_and_wait(err);
                continue 'main_loop;
            }
        };

        if books_vec.is_empty() {
            ui::display_formatted_text("Нет файлов к сбору.", Some(&red));
            thread::sleep(Duration::from_secs(SUCCESS_PAUSE_DURATION));
            continue 'main_loop;
        }
        
        ui::display_formatted_text("\nИдет анализ отобранных excel-файлов, ожидайте...", None);

        let acts_vec = {
            let mut temp_acts_vec = Vec::new();
            for item in books_vec.into_iter() {
                let book = match item {
                    Ok(x) => x,
                    Err(err) => {
                        let _ = Term::stdout().clear_last_lines(2); // удаляется сообщение что идет анализ excel-файлов
                        display_error_and_wait(err);
                        continue 'main_loop;
                    }
                };

                let sheet = match Sheet::new(book, &user_entered_sh_name) {
                    Ok(x) => x,
                    Err(err) => {
                        let _ = Term::stdout().clear_last_lines(2); // удаляется сообщение что идет анализ excel-файлов
                        display_error_and_wait(err);
                        continue 'main_loop;
                    }
                };

                let act = match Act::new(sheet) {
                    Ok(x) => x,
                    Err(err) => {
                        let _ = Term::stdout().clear_last_lines(2); // удаляется сообщение что идет анализ excel-файлов
                        display_error_and_wait(err);
                        continue 'main_loop;
                    }
                };

                temp_acts_vec.push(act);
            }
            temp_acts_vec
        };

        let _ = Term::stdout().clear_last_lines(2); // удаляется сообщение что идет анализ excel-файлов
        ui::display_formatted_text("\nИдет расчет структуры результирующего excel-отчета, ожидайте...", None);

        // "При вызове new() для Report требуется вектор актов. Это связанно с тем, что xlsxwriter
        // не может вставлять столбцы и не сможет переносить то, что им уже записано (т.к. не умеет читать Excel),
        // что предполагает необходимость установить общее количество столбцов, и их порядок до того как начнется запись.
        // Получается, на протяжении работы программы в Report акты передаются дважды:
        // при создании формы отчета для создания выборки всех названий, что встречаются в итогах,
        // а второй раз акт в Report будет передан циклом записи."

        let wrappedreport = Report::new(&report_path, &acts_vec);
        let _ = Term::stdout().clear_last_lines(2); // удаляется сообщение что идет вычисление структуры excel-отчета

        let mut report = match wrappedreport {
            Ok(rep) => rep,
            Err(err) => {
                display_error_and_wait(err);
                continue 'main_loop;
            }
        };

        ui::display_formatted_text("\nИдет запись результирующего excel-отчета, ожидайте...", None);

        for act in acts_vec.iter() {
            match report.write(act) {
                Ok(updated_report) => report = updated_report,
                Err(err) => {
                    let _ = Term::stdout().clear_last_lines(2); // удаляется сообщение что записывается Excel
                    display_error_and_wait(err);
                    continue 'main_loop;
                }
            }
        }

        let files_counter = report.body_syze_in_row;

        if let Err(err) = report.write_and_close_report(&report_path) {
            let _ = Term::stdout().clear_last_lines(2); // удаляется сообщение что записывается Excel
            display_error_and_wait(err);
            continue 'main_loop;
        }

        let _ = Term::stdout().clear_last_lines(2); // удаляется сообщение что записывается Excel
        ui::display_formatted_text("\nУспешно выполнено.", Some(&cyan));

        let base_msg = format!("Собрано {files_counter} файла(ов).");
        let footer_msg = format!(r#"Создан файл "{}""#, report_path.display());
        let full_msg = format!("{base_msg}\n{footer_msg}\n");

        ui::display_formatted_text(&full_msg, None);
        thread::sleep(Duration::from_secs(SUCCESS_PAUSE_DURATION));
        continue 'main_loop;
    }
}

fn display_error_and_wait(err: Error<'_>) {
    let red = Style::new().red();
    ui::display_formatted_text("\nВозникла ошибка.", Some(&red));

    let error_message = format!("\n{}", &err.to_string());
    ui::display_formatted_text(&error_message, None);

    // Подсчет количества не пустых и не состоящих только из пробелов строк в сообщении об ошибке
    let line_count = error_message
        .lines()
        .filter(|line| !line.trim().is_empty())
        .count();

    // Если строк меньше или равно 3, задержка 2 секунды, иначе 3 секунды
    let sleep_time_in_sec = if line_count <= 3 { 2 } else { 3 };
    thread::sleep(Duration::from_secs(sleep_time_in_sec));
}
