use crate::error::Error;
use calamine::{DataType, Range, Reader, Xlsx, XlsxError};
use std::collections::HashMap;
use std::fs::File;
use std::io::BufReader;
use std::path::PathBuf;
use walkdir::{DirEntry, WalkDir};

const EXCEL_FILE_EXTENSION: &str = ".xlsm";

#[derive(PartialEq)]
pub enum Required {
    Y,
    N,
}

// Магические цифры в кортеже это смещение в столбцах от первого столбца в документе до столбца, содержащего указанный литерал. Такую запись легче воспринимать.
// Литералы маленькими буквами из-за того что, например, "Доп. соглашение" Excel переведет в "Доп. Соглашение" если встать в ячейку и нажать Enter. Перестраховка от сюрпризов
pub const SEARCH_REFERENCE_POINTS: [(usize, Required, &str); 10] = [
    (0, Required::N, "исполнитель"),
    (0, Required::Y, "стройка"),
    (0, Required::Y, "объект"),
    (9, Required::Y, "договор подряда"),
    (9, Required::Y, "доп. соглашение"),
    (5, Required::Y, "номер документа"),
    (3, Required::Y, "наименование работ и затрат"),
    (11, Required::N, "зтр всего чел.-час"),
    (0, Required::N, "итого по акту:"),
    (3, Required::Y, "стоимость материальных ресурсов (всего)"),
];

#[derive(Debug, Clone)]
pub struct DesiredData {
    pub name: &'static str,
    pub offset: Option<(&'static str, (i8, i8))>,
}
#[rustfmt::skip]
pub const DESIRED_DATA_ARRAY: [DesiredData; 16] = [
    DesiredData{name:"Исполнитель",                  offset: None},
    DesiredData{name:"Глава",                        offset: None},
    DesiredData{name:"Глава наименование",           offset: None},
    DesiredData{name:"Объект",                       offset: Some(("объект",                         (0, 3)))},
    DesiredData{name:"Договор №",                    offset: Some(("договор подряда",                (0, 2)))},
    DesiredData{name:"Договор дата",                 offset: Some(("договор подряда",                (1, 2)))},
    DesiredData{name:"Смета №",                      offset: Some(("договор подряда",                (0, -9)))},
    DesiredData{name:"Смета наименование",           offset: Some(("договор подряда",                (1, -9)))},
    DesiredData{name:"По смете в ц.2000г.",          offset: Some(("доп. соглашение",                (0, -4)))},
    DesiredData{name:"Выполнение работ в ц.2000г.",  offset: Some(("доп. соглашение",                (1, -4)))},
    DesiredData{name:"Акт №",                        offset: Some(("номер документа",                (2, 0)))},
    DesiredData{name:"Акт дата",                     offset: Some(("номер документа",                (2, 4)))},
    DesiredData{name:"Отчетный период начало",       offset: Some(("номер документа",                (2, 5)))},
    DesiredData{name:"Отчетный период окончание",    offset: Some(("номер документа",                (2, 6)))},
    DesiredData{name:"Метод расчета",                offset: Some(("наименование работ и затрат",    (-1, -3)))},
    DesiredData{name:"Затраты труда, чел.-час",      offset: None},
];
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

pub struct Sheet {
    pub path: PathBuf,
    pub sheet_name: String,
    pub data: Range<DataType>,
    pub search_points: HashMap<&'static str, (usize, usize)>,
    pub range_start: (usize, usize),
}

impl<'a> Sheet {
    pub fn new(
        // разработчики Calamine делают зачем-то &mut self в функции worksheet_range(&mut self, name: &str),
        // из-за этого workbook приходится держать мутабельным, хотя этот код его менять вовсе не собирается
        // (из-за мутабельности workbook проблема при попытке множественных ссылок: можно только клонировать)
        workbook: &'a mut Book,
        user_entered_sh_name: &'a str,
        search_reference_points: &[(usize, Required, &'static str)],
        expected_sum_of_requir_col: usize,
    ) -> Result<Sheet, Error<'a>> {
        let sheet_name = workbook
            .data
            .sheet_names()
            .iter()
            .find(|name| name.to_lowercase() == user_entered_sh_name)
            .ok_or(Error::CalamineSheetOfTheBookIsUndetectable {
                sh_name_for_search: user_entered_sh_name,
                sh_names: workbook.data.sheet_names().to_owned(),
            })?
            .clone();

        let sheetXL = workbook
            .data
            .worksheet_range(&sheet_name)
            .ok_or(Error::CalamineSheetOfTheBookIsUndetectable {
                sh_name_for_search: user_entered_sh_name,
                sh_names: workbook.data.sheet_names().to_owned(),
            })?
            .or_else(|error| {
                Err(Error::CalamineSheetOfTheBookIsUnreadable {
                    sh_name: sheet_name.to_owned(),
                    err: error,
                })
            })?;

        let mut search_points = HashMap::new();

        let mut temp_sh_iter = sheetXL.used_cells();
        let mut temp;
        for item in search_reference_points {
            match item.1 {
                // Для Y-типов подходит расходуемый итератор - достигается проверка по очередности вохождения слов по строкам
                // (т.е. "Стройку" мы ожидаем выше "Объекта, например")
                Required::Y => {
                    temp = temp_sh_iter.find(|x| {
                        x.2.get_string()
                            .as_ref()
                            .unwrap_or_else(|| &"")
                            .to_lowercase()
                            == item.2
                    });
                }
                // Для N-типов нельзя использовать расходуемые итераторы, так как необязательное значение может и отсутсвовать (и при его поиске израсходуется итератор)
                Required::N => {
                    temp = sheetXL.used_cells().find(|x| {
                        x.2.get_string()
                            .as_ref()
                            .unwrap_or_else(|| &"")
                            .to_lowercase()
                            == item.2
                    });
                }
            }

            if let Some((row, col, _)) = temp {
                search_points.insert(item.2, (row, col));
            }
        }

        // Проверка на полноту данных
        let test = SEARCH_REFERENCE_POINTS
            .iter()
            .filter(|x| x.1 == Required::Y)
            .last()
            .unwrap_or_else(|| panic!("ложь: \"DESIRED_DATA_ARRAY всегда имеет значения\""))
            .2;

        search_points
            .get(test)
            .ok_or(Error::SheetNotContainAllNecessaryData)?;

        // Проверка значений на удаленность столбцов, чтобы гарантировать что найден нужный лист.
        let first_col = search_points
            .get("стройка")
            .unwrap_or_else(|| panic!("ложь: \"Необеспечены действительные имена HashMap\""));

        let (just_a_amount_requir_col, just_a_sum_requir_col) = search_reference_points
            .iter()
            .fold((0_usize, 0), |acc, item| match item.1 {
                Required::Y => (
                    acc.0 + 1,
                    acc.1
                        + search_points
                            .get(item.2)
                            .unwrap_or_else(|| {
                                panic!("ложь: \"Необеспечены действительные имена HashMap\"")
                            })
                            .1,
                ),
                _ => acc,
            });

        if let false = just_a_sum_requir_col - first_col.1 * just_a_amount_requir_col
            == expected_sum_of_requir_col
        {
            return Err(Error::ShiftedColumnsInHeader);
        }
        let range_start_u32 = sheetXL
            .start()
            .ok_or(Error::EmptySheetRange(user_entered_sh_name))?;

        let range_start = (range_start_u32.0 as usize, range_start_u32.1 as usize);

        Ok(Sheet {
            path: workbook.path.clone(),
            sheet_name,
            data: sheetXL,
            search_points,
            range_start,
        })
    }
}

pub fn get_vector_of_books(path: PathBuf) -> Result<Vec<Result<Book, XlsxError>>, Error<'static>> {
    let books_vec = match path.is_dir() {
        true => {
            let temp_res = directory_traversal(&path);
            let books_vector_len = (temp_res.0).len();
            if books_vector_len == 0 {
                return Err(Error::NoFilesInSpecifiedPath(path));
            } else {
                println!(
                    "\n Обнаружено {} файлов с расширением \"{EXCEL_FILE_EXTENSION}\".",
                    books_vector_len + temp_res.1 as usize
                );
                if temp_res.1 > 0 {
                    println!(" Из них {} помечены \"@\" для исключения.", temp_res.1);
                } else {
                    println!(" Среди них нет файлов, помеченных как исключенные.");
                }
                println!("\n Идет отбор нужных файлов, ожидайте...");
            }
            Ok(temp_res.0)
        }
        false if path.is_file() => Ok(vec![Book::new(path)]),
        _ => panic!(" Введенный пользователем путь не является папкой или файлом"),
    };
    books_vec
}

fn directory_traversal(path: &PathBuf) -> (Vec<Result<Book, XlsxError>>, u32) {
    let prefix = path.to_string_lossy().to_string();
    let parent = path.parent().unwrap().to_string_lossy().to_string();

    let is_excluded_file = |entry: &DirEntry| -> bool {
        entry
            .path()
            .strip_prefix(&prefix)
            .unwrap()
            .to_string_lossy()
            .contains('@')
    };

    let walkdir = WalkDir::new(path)
        .into_iter()
        .filter_map(|e| e.ok()) //будет молча пропускать каталоги, на доступ к которым у владельца запущенного процесса нет разрешения
        .filter(|e| {
            e.file_name()
                .to_str()
                .map(|s| !s.starts_with('~') & s.ends_with(EXCEL_FILE_EXTENSION))
                .unwrap_or_else(|| false)
        });

    let mut counter = 1;
    let mut books_vector = vec![];
    let mut excluded_files_counter = 0_u32;

    for entry in walkdir {
        if is_excluded_file(&entry) {
            excluded_files_counter += 1;
            continue;
        }

        let path_to_processed_file = entry
            .path()
            .strip_prefix(&parent)
            .unwrap()
            .to_string_lossy()
            .to_string();

        let temp_book = Book::new(entry.into_path());
        books_vector.push(temp_book);

        println!(" {}: {}", counter, path_to_processed_file);
        counter += 1;
    }
    (books_vector, excluded_files_counter)
}
