use calamine::{DataType, Range, Xlsx, XlsxError};
use std::collections::HashMap;
use std::fs::File;
use std::io::BufReader;

pub enum Required {
    Y,
    N,
}

pub const SEARCH_REFERENCE_POINTS: [(usize, Required, &str); 8] = [
    (0, Required::N, "Исполнитель"),
    (0, Required::Y, "Стройка"),
    (0, Required::Y, "Объект"),
    (9, Required::Y, "Договор подряда"),
    (9, Required::N, "Доп. соглашение"),
    (5, Required::Y, "Номер документа"),
    (3, Required::Y, "Наименование работ и затрат"),
    (3, Required::Y, "Стоимость материальных ресурсов (всего)"),
];

#[rustfmt::skip]
pub const NAMES_OF_HEADER: [(&'static str, Option<(&'static str, (i8, i8))>); 15] = [
    ("Исполнитель", None),
    ("Глава", None),
    ("Глава наименование", None),
    ("Объект", Some(("Объект", (0, 3)))),
    ("Договор №", Some(("Договор подряда", (0, 2)))),
    ("Договор дата", Some(("Договор подряда", (1, 2)))),
    ("Смета №", Some(("Договор подряда", (0, -9)))),
    ("Смета наименование", Some(("Договор подряда", (1, -9)))),
    ("По смете в ц.2000г.", Some(("Договор подряда", (2, -4)))),
    ("Выполнение работ в ц.2000г.", Some(("Договор подряда", (3, -4)))),
    ("Акт №", Some(("Номер документа", (2, 0)))),
    ("Акт дата", Some(("Номер документа", (2, 4)))),
    ("Отчетный период начало", Some(("Номер документа", (2, 5)))),
    ("Отчетный период окончание", Some(("Номер документа", (2, 6)))),
    ("Метод расчета", Some(("Наименование работ и затрат", (-1, -3)))),
];
pub struct Book {
    pub path: String,
    pub data: Xlsx<BufReader<File>>,
}

impl Book {
    pub fn new(path: &str) -> Result<Self, XlsxError> {
        let data: Xlsx<_> = calamine::open_workbook(&path)?;
        Ok(Book {
            path: path.to_owned(),
            data,
        })
    }
}

pub struct Sheet {
    pub path: String,
    pub sheetname: &'static str,
    pub data: Range<DataType>,
    pub search_points: HashMap<&'static str, (usize, usize)>,
}
