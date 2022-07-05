use console::Term;
use std::env;
use std::path::Path;
use walkdir::{DirEntry, WalkDir};
use xlsxwriter::Workbook;
mod extract;
mod load;
mod transform;
mod ui;
use crate::extract::{Book, Sheet, SEARCH_REFERENCE_POINTS};
use crate::load::Report;
use crate::transform::Act;

fn main() {
    // let (path_str, sh_name) = ui::session();
    let sh_name = "лист1".to_owned();
    let path_str = r"C:\Users\User\rust\ks2_etl";
    let report_name_prefix = env::args()
        .next()
        .unwrap()
        .trim_end_matches(".exe")
        .to_owned();

    let report_name = report_name_prefix + ".xlsx";
    let wb = Workbook::new(&report_name);
    let mut report = Report::new(wb);

    let path = Path::new(&path_str);

    let is_excluded_file = |entry: &DirEntry| -> bool {
        entry
            .path()
            .strip_prefix(&path_str)
            .unwrap()
            .to_string_lossy()
            .contains('@')
    };

    let mut excluded_files_counter = 0_u32;
    for entry in WalkDir::new(path)
        .into_iter()
        .filter_map(|e| e.ok()) //будет молча пропускать каталоги, на доступ к которым у владельца запущенного процесса нет разрешения
        .filter(|e| {
            e.file_name()
                .to_str()
                .map(|s| !s.starts_with('~') & s.ends_with(".xlsm"))
                .unwrap_or(false)
        })
    {
        if is_excluded_file(&entry) {
            excluded_files_counter += 1;
            continue;
        }

        let file_path = entry.path();

        let mut file = Book::new(file_path.to_path_buf()).unwrap();

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
                println!("\nФайл, вызывающий ошибку: {}", file_path.display());
                loop {}
            };
            panic!()
        });

        let act = Act::new(sheet).unwrap_or_else(|err| {
            let _ = Term::stdout().clear_screen();
            println!("\nВозникла ошибка. \n{}", err);
            println!("\nФайл, вызывающий ошибку: {}", file_path.display());
            loop {}
        });

        if let Err(err) = report.write(&act) {
            let _ = Term::stdout().clear_screen();
            println!("\nВозникла ошибка. \n{}", err);
            println!("\nФайл, вызывающий ошибку: {}", file_path.display());
            loop {}
        };
        let _ = Term::stdout().clear_screen();
        println!(
            "Успешно собрана информация из {} актов",
            report.empty_row - 1
        );
    }
    let files_counter = report.empty_row - 1;
    let file = report.end();
    file.unwrap().close().unwrap_or_else(|_| {
        let _ = Term::stdout().clear_screen();
        println!(
            "Возникла ошибка, вероятная причина:\
        \nне закрыт файл Excel с результатами прошлого сбора."
        );
    });

    let _ = Term::stdout().clear_screen();
    println!("Успешно выполнено.\nСобрано {} файла(ов).", files_counter);
    if excluded_files_counter > 0 {
        println!("{} файла(ов), помеченно «@», для исключения.", excluded_files_counter);
    }
    println!("\nСоздан файл \"{}\"", report_name);
    loop {}
}
