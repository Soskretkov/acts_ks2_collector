use console::Term;
use glob::{glob_with, MatchOptions};
use std::env;
use xlsxwriter::Workbook;
#[macro_use]
extern crate error_chain;
// extern crate walkdir;

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
        Glob(glob::GlobError);
        Pattern(glob::PatternError);
    }
}

fn main() -> Result<()> {
    // let (path, sh_name) = ui::session();
    let sh_name = "лист1".to_owned();
    let path = r"C:\Users\User\rust\acts_ks2_etl\";
    let report_name_prefix = env::args()
        .next()
        .unwrap()
        .trim_end_matches(".exe")
        .to_owned();

    let report_name = report_name_prefix + ".xlsx";
    let wb = Workbook::new(&report_name);
    let mut report = Report::new(wb);
    let options = MatchOptions {
        case_sensitive: false,
        ..Default::default()
    };

    for entry in glob_with(r"**/*/[@]*.xlsm", options)? {
        // let f_path = entry.unwrap().as_path().to_string_lossy().to_string();
        let f_path = entry?.to_string_lossy().to_string();

        let mut file = Book::new(&f_path).unwrap();

        let sheet = Sheet::new(
            &mut file,
            &sh_name,
            &SEARCH_REFERENCE_POINTS,
            29, //передается для расчета смещения столбцов. Это сумма номеров столбцов Y-типа в DESIRED_DATA_ARRAY: 0 + 0 + 3 + 5 + 9 + 9 + 3.
        )
        .unwrap_or_else(|err| {
            if let Some(text) = ks2_etl::error_message(err, &sh_name) {
                let _ = Term::stdout().clear_screen();
                println!("\nВозникла ошибка. \n{}", text);
                println!("\nФайл, вызывающий ошибку: {}", f_path);
                loop {}
            };
            panic!()
        });

        let act = Act::new(sheet).unwrap_or_else(|err| {
            let _ = Term::stdout().clear_screen();
            println!("\nВозникла ошибка. \n{}", err);
            println!("\nФайл, вызывающий ошибку: {}", f_path);
            loop {}
        });

        if let Err(err) = report.write(&act) {
            let _ = Term::stdout().clear_screen();
            println!("\nВозникла ошибка. \n{}", err);
            println!("\nФайл, вызывающий ошибку: {}", f_path);
            loop {}
        };
        println!(
            "Успешно собрана информация из {} актов",
            report.empty_row - 1
        );
    }
    let number_of_files = report.empty_row - 1;
    let file = report.end();
    file.unwrap().close().unwrap_or_else(|_| {
        let _ = Term::stdout().clear_screen();
        println!(
            "Возникла ошибка, вероятная причина:\
        \nне закрыт файл Excel с результатами прошлого сбора."
        );
    });

    let _ = Term::stdout().clear_screen();
    println!(
        "Успешно выполнено.\nСобрано {} файлов.\nСоздан файл \"{}\"",
        number_of_files, report_name
    );
    loop {}
    // Ok(())
}
