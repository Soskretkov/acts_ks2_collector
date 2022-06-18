use calamine::{DataType, Range, Reader, Xlsx, XlsxError};
use std::collections::HashMap;
use std::fs::File;
use std::io::BufReader;

#[derive(PartialEq)]
pub enum Required {
    Y,
    N,
}

// Маленькими буквами из-за того что, например, "Доп. соглашение" Excel переведет в "Доп. Соглашение" если встать в ячейку и нажать Enter. Перестраховка от сюрпризов
pub const SEARCH_REFERENCE_POINTS: [(usize, Required, &str); 8] = [
    (0, Required::N, "исполнитель"),
    (0, Required::Y, "стройка"),
    (0, Required::Y, "объект"),
    (9, Required::Y, "договор подряда"),
    (9, Required::Y, "доп. соглашение"),
    (5, Required::Y, "номер документа"),
    (3, Required::Y, "наименование работ и затрат"),
    (3, Required::Y, "стоимость материальных ресурсов (всего)"),
];

#[derive(Debug, Clone)]
pub struct DesiredData {
    pub name: &'static str,
    pub offset: Option<(&'static str, (i8, i8))>,
}
#[rustfmt::skip]
pub const DESIRED_DATA_ARRAY: [DesiredData; 15] = [
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
    pub range_start: (usize, usize),
}

impl Sheet {
    pub fn new<'a>(
        workbook: &'a mut Book,
        sheetname: &'static str,
        search_reference_points: &[(usize, Required, &'static str)],
        expected_sum_of_requir_col: usize,
    ) -> Result<Sheet, &'static str> {
        let data = match workbook.data.worksheet_range(sheetname) {
            Some(x) => x.unwrap(),
            None => return Err("Отсуствует запрошенный лист"),
        };

        let mut search_points = HashMap::new();

        let mut temp_sh_iter = data.used_cells();
        let mut temp;
        for item in search_reference_points {
            match item.1 {
                // Для Y-типов расходуемый итератор - тем самым достигается проверка по очередности вохождения слов по строкам
                // (т.е. "Стройку" мы ожидаем выше "Объекта, например")
                Required::Y => {
                    temp = temp_sh_iter.find(|x| {
                        x.2.get_string().as_ref().unwrap_or(&"").to_lowercase() == item.2
                    });
                }
                // Для N-типов нельзя использовать расходуемые итераторы, так как необязательное значение будет отсутсвовать (и при его поиске израсходуется итератор)
                Required::N => {
                    temp = data.used_cells().find(|x| {
                        x.2.get_string().as_ref().unwrap_or(&"").to_lowercase() == item.2
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
            .unwrap() //unwrap не требует обработки: SEARCH_REFERENCE_POINTS всегда имеет значения
            .2;
        if search_points.get(test).is_none() {
            return Err("Лист не содержит всех необходимых данных");
        }

        // Проверка значений на удаленность столбцов, чтобы гарантировать что найден нужный лист.
        let first_col = search_points.get("стройка").unwrap().1;

        let (just_a_amount_requir_col, just_a_sum_requir_col) = search_reference_points
            .iter()
            .fold((0_usize, 0), |acc, item| match item.1 {
                Required::Y => (acc.0 + 1, acc.1 + search_points.get(item.2).unwrap().1),
                _ => acc,
            });

        if let false = just_a_sum_requir_col - first_col * just_a_amount_requir_col
            == expected_sum_of_requir_col
        {
            return Err("Нетипичное заглавие (шапка) КС-2");
        }

        let range_start = match data.start() {
            Some(x) => (x.0 as usize, x.1 as usize),
            None => return Err("Не удалось получить начало диапазона листа"),
        };

        Ok(Sheet {
            path: workbook.path.clone(),
            sheetname,
            data,
            search_points,
            range_start,
        })
    }
}
