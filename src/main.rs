use xlsxwriter::{Workbook};
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
    // let (path, sh_name) = ui::session();
    let sh_name = "Лист1";
    let path = r"C:\Users\User\rust\acts_ks2_etl\";
    fn is_not_temp(entry: &DirEntry) -> bool {
        entry
            .file_name()
            .to_str()
            .map(|s| !s.starts_with('~') && !s.contains('@'))
            .unwrap_or(false)
    }

    let wb = Workbook::new("Test.xlsx");
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
            .unwrap();

            let act = Act::new(sheet).unwrap();

            if let Err(x) = report.write(&act) {
                println!("{x}");
            };
        }
    }
    let wb_2 = report.end();
    let _ = wb_2.unwrap().close();
    Ok(())
}
