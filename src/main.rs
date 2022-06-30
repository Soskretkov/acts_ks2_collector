use std::env;
use xlsxwriter::Workbook;
#[macro_use]
extern crate error_chain;
extern crate walkdir;
use walkdir::{DirEntry, WalkDir};
mod extract;
mod load;
mod transform;
mod ui;
use crate::extract::{Book, Sheet, SEARCH_REFERENCE_POINTS};
use crate::load::Report;
use crate::transform::Act;

error_chain! {
    foreign_links {
        WalkDir(walkdir::Error);
        Io(std::io::Error);
    }
}

fn main() -> Result<()> {
    let (path, sh_name) = ui::session();
    // let sh_name = "Лист1";
    // let path = r"C:\Users\User\rust\acts_ks2_etl\";
    fn is_not_temp(entry: &DirEntry) -> bool {
        entry
            .file_name()
            .to_str()
            .map(|s| !s.starts_with('~') && !s.contains('@'))
            .unwrap_or(false)
    }
    let report_name_prefix = env::args()
        .next()
        .unwrap()
        .strip_suffix(".exe")
        .unwrap()
        .to_owned();
    let report_name = report_name_prefix + ".xlsx";
    let wb = Workbook::new(&report_name);
    let mut report = Report::new(wb);

    for entry in WalkDir::new(path)
        .into_iter()
        .filter_map(|e| e.ok()) //будет молча пропускать каталоги, на доступ к которым у владельца запущенного процесса нет разрешения
        .filter(is_not_temp)
    {
        let f_path = entry.path().to_string_lossy();

        if f_path.ends_with(".xlsm") {
            let mut file = Book::new(&f_path.to_string()).unwrap();

            let sheet = Sheet::new(
                &mut file,
                &sh_name,
                &SEARCH_REFERENCE_POINTS,
                29, //передается для расчета смещения столбцов. Это сумма номеров столбцов Y-типа в DESIRED_DATA_ARRAY: 0 + 0 + 3 + 5 + 9 + 9 + 3.
            )
            .unwrap_or_else(|err| {
                if let Some(text) = acts_ks2_etl::error_message(err, &sh_name) {
                    println!("Возникла ошибка. \n{}", text);
                    println!("\nФайл, вызывающий ошибку: {}", f_path.to_string());
                    loop {}
                };
                panic!()
            });

            let act = Act::new(sheet).unwrap_or_else(|err| {
                println!("Возникла ошибка. \n{}", err);
                println!("\nФайл, вызывающий ошибку: {}", f_path.to_string());
                loop {}
            });

            if let Err(err) = report.write(&act) {
                println!("Возникла ошибка. \n{}", err);
                println!("\nФайл, вызывающий ошибку: {}", f_path.to_string());
                loop {}
            };
        }
    }
    let number_of_files = report.empty_row - 1;
    let file = report.end();
    file.unwrap().close().unwrap_or_else(|_| {
        println!(
            "Возникла ошибка, вероятная причина:\
        \nне закрыт файл Excel с результатами прошлого сбора."
        );
    });

    println!(
        "Успешно выполнено.\nСобрано {} файлов.\nСоздан файл \"{}\"",
        number_of_files, report_name
    );
    loop {}
    // Ok(())
}
