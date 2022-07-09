use console::Term; // для очистки консоли перед выводом полезных сообщений
use std::env; // имя ".exe" будет присвоено файлу Excel
use std::thread; // для засыпания на секунду-две при печати сообщений
use std::time::Duration; // для засыпания на секунду-две при печати сообщений
use xlsxwriter::Workbook;
mod extract;
mod load;
mod transform;
mod ui;
use crate::extract::{Book, Sheet, SEARCH_REFERENCE_POINTS};
use crate::load::Report;
use crate::transform::Act;

fn main() {
    Term::stdout().set_title("Ks2 etl");
    ui::show_first_lines();
    ui::show_help();
    'main_loop: loop {
        // let (path, sh_name) = ui::user_input();
        let sh_name = "Лист1".to_owned();
        let path = std::path::PathBuf::from(r"C:\Users\User\rust\ks2_etl".to_string());

        let sh_name_lowercase = sh_name.to_lowercase();

        let report_name_prefix = env::args()
            .next()
            .unwrap()
            .trim_end_matches(".exe")
            .to_owned();

        let report_name = report_name_prefix + ".xlsx";
        let wb = Workbook::new(&report_name);
        let mut report = Report::new(wb);

        let _ = Term::stdout().clear_screen();
        let (books_vector, excluded_files_counter) = match path.is_dir() {
            true => {
                let temp_res = extract::directory_traversal(&path);
                let books_vector_len = (temp_res.0).len();
                if books_vector_len == 0 {
                    println!(
                        "Нет файлов \".xlsm\" по указанному пути: {}",
                        path.display()
                    );
                    thread::sleep(Duration::from_secs(1));
                    continue 'main_loop;
                } else {
                    println!(
                        "Обнаружено {} файлов с расширением \".xlsm\".",
                        books_vector_len + temp_res.1 as usize
                    );
                    if temp_res.1 > 0 {
                        println!("Из них {} помечены \"@\" для исключения.", temp_res.1);
                    } else {
                        println!("Среди них нет файлов, помеченных как исключенные.");
                    }
                    println!("\nИдет обработка, ожидайте...\n");
                }
                temp_res
            }
            false if path.is_file() => (vec![Book::new(path)], 0_u32),
            _ => panic!("Введенный пользователем путь не является папкой или файлом"),
        };

        for mut item in books_vector.into_iter() {
            let book = item.as_mut().unwrap();
            let wrapped_sheet = Sheet::new(
                book,
                &sh_name_lowercase,
                &SEARCH_REFERENCE_POINTS,
                29, //передается для расчета смещения столбцов. Это сумма номеров столбцов Y-типа в DESIRED_DATA_ARRAY: 0 + 0 + 3 + 5 + 9 + 9 + 3.
            );

            let sheet = match wrapped_sheet {
                Ok(x) => x,
                Err(err) => {
                    if let Some(text) = ks2_etl::error_message(err, &sh_name) {
                        println!("\nВозникла ошибка.\n{}", text);
                        println!("\nФайл, вызывающий ошибку: {}", book.path.display());
                        thread::sleep(Duration::from_secs(3));
                        println!("\n\n\n\n");
                        continue 'main_loop;
                    };
                    panic!()
                }
            };

            let wrapped_act = Act::new(sheet);
            let act = match wrapped_act {
                Ok(x) => x,
                Err(err) => {
                    println!("\nВозникла ошибка.\n{}", err);
                    println!("\nФайл, вызывающий ошибку: {}", book.path.display());
                    thread::sleep(Duration::from_secs(3));
                    println!("\n\n\n\n");
                    continue 'main_loop;
                }
            };

            if let Err(err) = report.write(&act) {
                println!("\nВозникла ошибка.\n{}", err);
                println!("\nФайл, вызывающий ошибку: {}", book.path.display());
                thread::sleep(Duration::from_secs(3));
                println!("\n\n\n\n");
                continue 'main_loop;
            };
        }
        let files_counter = report.empty_row - 1;
        if report.end().unwrap().close().is_err() {
            println!(
                    "Возникла ошибка, вероятная причина:\nне закрыт файл Excel с результатами прошлого сбора."
                );
            thread::sleep(Duration::from_secs(3));
            println!("\n\n\n\n");
            continue 'main_loop;
        }

        println!("Успешно выполнено.\nСобрано {} файла(ов).", files_counter);
        if excluded_files_counter > 0 {
            println!(
                "{} файла(ов) не были собраны, поскольку они помечены \"@\" для исключения.",
                excluded_files_counter
            );
        }
        println!("\nСоздан файл \"{}\"", report_name);
        thread::sleep(Duration::from_secs(1));
        println!("\n\n");
        continue 'main_loop;
    }
}
