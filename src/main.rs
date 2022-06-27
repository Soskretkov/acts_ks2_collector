use xlsxwriter::Workbook;
#[macro_use]
extern crate error_chain;
extern crate walkdir;
mod extract;
mod load;
mod transform;
mod ui;
use crate::extract::{Book, Sheet, SEARCH_REFERENCE_POINTS};
use crate::load::Report;
use crate::transform::Act;

use walkdir::{DirEntry, WalkDir};

error_chain! {
    foreign_links {
        WalkDir(walkdir::Error);
        Io(std::io::Error);
    }
}

fn main() -> Result<()> {
    // let (_path, _sh_name) = ui::session();

    // let path = "f.xlsm";
    let path = r"C:\Users\User\rust\acts_ks2_etl\";
    fn is_not_temp(entry: &DirEntry) -> bool {
        entry
            .file_name()
            .to_str()
            .map(|s| !s.starts_with("~") && !s.contains('@'))
            .unwrap_or(false)
    }

    for entry in WalkDir::new(path)
        .into_iter()
        .filter_map(|e| e.ok()) //будет молча пропускать каталоги, на доступ к которым у владельца запущенного процесса нет разрешения
        .filter(|f| is_not_temp(f))
    {
        let f_path = entry.path().to_string_lossy();

        if f_path.ends_with(".xlsm") {
            let mut file = Book::new(&f_path.to_string()).unwrap();

            let sheet = Sheet::new(
                &mut file,
                "Лист1",
                &SEARCH_REFERENCE_POINTS,
                29, //передается для расчета смещения столбцов. Это сумма номеров столбцов Y-типа в DESIRED_DATA_ARRAY: 0 + 0 + 3 + 5 + 9 + 9 + 3.
            )
            .unwrap();

            let act = Act::new(sheet).unwrap();

    let wb = Workbook::new("Test.xlsx");
            let mut report = Report::new(wb);
            if let Err(x) = report.write(&act) {
                println!("{x}");
            };
            let wb_2 = report.end();
            let _ = wb_2.unwrap().close();
        }
    }
    // println!("{:#?}", report.part_1_just);
    // println!("{:#?}", report.part_2_base);
    // println!("{:#?}", report.part_3_curr);

    // println!("{:#?}", vector_of_acts[0].data_of_totals);

    // let (_part_2_fst_fls_base, _part_4_fst_fls_curr) = load::first_file_data_names(&vector_of_acts[0].data_of_totals);
    //   println!("{:#?}", _part_2_fst_fls_base);
    //   println!("{:#?}", _part_4_fst_fls_curr);

    // Печать шапки
    // let mut header = act
    //     .names_of_header
    //     .iter()
    //     .zip(act.data_of_header.iter().map(|x| {
    //         x.as_ref()
    //             .unwrap_or(&transform::DateVariant::String("".to_string()))
    //             .clone()
    //     }));

    // for print in header {
    //      println!("{}:  {:?}", print.0.content, print.1);
    // }
    // let print = header.nth(2).unwrap();
    // println!("{}:  {:?}", print.0.name, print.1);

    // // Печать итогов
    // let summary = act.data_of_totals;
    // let last_row = summary.iter().last().unwrap();
    // println!(
    //     "\n{}: базовая - {}; текущая - {}",
    //     last_row.name,
    //     last_row.base_price[0].unwrap_or(0.),
    //     last_row.current_price[0].unwrap_or(0.)
    // );

    // Печать rows Excel
    // let sheet = Sheet::new(
    //     &mut wb,
    //     "Лист1",
    //     &SEARCH_REFERENCE_POINTS,
    //     0 + 0 + 3 + 5 + 9 + 9 + 3,
    // )
    // .unwrap();

    // let start_of_range = sheet.search_points
    //     .get("стоимость материальных ресурсов (всего)")
    //     .unwrap();
    // for row in sheet.data.rows().skip(start_of_range.0) {
    //     println!("{:?}", row);
    // }
    Ok(())
}
