use crate::errors::Error;
use std::collections::HashMap;

#[derive(Clone, Copy)]
pub enum Column {
    Initial,
    Contract,
}

#[derive(Clone, Copy)]
pub enum Row {
    TableHeader,
}

#[derive(Debug, Clone, PartialEq, Hash, Eq, Copy)]
pub enum TagID {
    Исполнитель,
    Стройка,
    Объект,
    ДоговорПодряда,
    ДопСоглашение,
    НомерДокумента,
    НаименованиеРаботИЗатрат,
    СтоимостьВЦенах2001,
    СтоимостьВТекущихЦенах,
    ЗтрВсего,
    ИтогоПоАкту,
    СтоимостьМатериальныхРесурсовВсего,
}

#[rustfmt::skip]
impl TagID {
    pub fn as_str(&self) -> &'static str {
        match self {
            TagID::Исполнитель => "Исполнитель",
            TagID::Стройка => "Стройка",
            TagID::Объект => "Объект",
            TagID::ДоговорПодряда => "Договор подряда",
            TagID::ДопСоглашение => "Доп. соглашение",
            TagID::НомерДокумента => "Номер документа",
            TagID::НаименованиеРаботИЗатрат => "Наименование работ и затрат",
            TagID::СтоимостьВЦенах2001 => "Стоимость в ценах 2001",
            TagID::СтоимостьВТекущихЦенах => "Стоимость в текущих ценах",
            TagID::ЗтрВсего => "ЗТР всего чел",
            TagID::ИтогоПоАкту => "Итого по акту:",
            TagID::СтоимостьМатериальныхРесурсовВсего => "Стоимость материальных ресурсов (всего)",
        }
    }
}

// режим сравнения двух текстов: частичное или полное совпадение
#[derive(Clone, Copy)]
pub enum TextCmp {
    Part,
    Whole,
}

#[derive(Clone, Copy)]
pub struct TagInfo {
    pub id: TagID,
    pub is_required: bool,
    pub group_by_row: Option<Row>,
    pub group_by_col: Option<Column>,
    pub look_at: TextCmp,
    pub match_case: bool,
}

// Перечислены в порядке вхождения на листе Excel при чтении ячеек слева направо и сверху вниз  (вхождение по строкам важно для валидации)
// Группировка по строке и столбцу потребуется для валидации в будующих версиях программы (не реализовано)
#[rustfmt::skip]
pub const TAG_INFO_ARRAY: [TagInfo; 12] = [
    TagInfo { id: TagID::Исполнитель,                        is_required: false, group_by_row: None,                   group_by_col: Some(Column::Initial),  look_at: TextCmp::Whole, match_case: false },
    TagInfo { id: TagID::Стройка,                            is_required: true,  group_by_row: None,                   group_by_col: Some(Column::Initial),  look_at: TextCmp::Whole, match_case: false },
    TagInfo { id: TagID::Объект,                             is_required: true,  group_by_row: None,                   group_by_col: Some(Column::Initial),  look_at: TextCmp::Whole, match_case: false },
    TagInfo { id: TagID::ДоговорПодряда,                     is_required: true,  group_by_row: None,                   group_by_col: Some(Column::Contract), look_at: TextCmp::Whole, match_case: false },
    TagInfo { id: TagID::ДопСоглашение,                      is_required: true,  group_by_row: None,                   group_by_col: Some(Column::Contract), look_at: TextCmp::Whole, match_case: false },
    TagInfo { id: TagID::НомерДокумента,                     is_required: true,  group_by_row: None,                   group_by_col: None,                   look_at: TextCmp::Whole, match_case: false },
    TagInfo { id: TagID::НаименованиеРаботИЗатрат,           is_required: true,  group_by_row: Some(Row::TableHeader), group_by_col: None,                   look_at: TextCmp::Whole, match_case: false },
    TagInfo { id: TagID::СтоимостьВЦенах2001,                is_required: true,  group_by_row: Some(Row::TableHeader), group_by_col: None,                   look_at: TextCmp::Part,  match_case: true },
    TagInfo { id: TagID::СтоимостьВТекущихЦенах,             is_required: true,  group_by_row: Some(Row::TableHeader), group_by_col: None,                   look_at: TextCmp::Part,  match_case: true },
    TagInfo { id: TagID::ЗтрВсего,                           is_required: false, group_by_row: Some(Row::TableHeader), group_by_col: None,                   look_at: TextCmp::Part,  match_case: true },
    TagInfo { id: TagID::ИтогоПоАкту,                        is_required: false, group_by_row: None,                   group_by_col: Some(Column::Initial),  look_at: TextCmp::Whole, match_case: true },
    TagInfo { id: TagID::СтоимостьМатериальныхРесурсовВсего, is_required: true,  group_by_row: None,                   group_by_col: None,                   look_at: TextCmp::Whole, match_case: false },
];

pub struct TagArrayTools;

impl TagArrayTools {
    pub fn _get_tags() -> &'static [TagInfo] {
        &TAG_INFO_ARRAY
    }
    pub fn get_tag_info_by_id(id: TagID) -> Result<TagInfo, Error<'static>> {
        TAG_INFO_ARRAY
            .into_iter()
            .find(|tag_info| tag_info.id == id)
            .ok_or_else(|| Error::InternalLogic {
                tech_descr: format!(r#"Массив тегов не содержит тег "{}""#, id.as_str()),
                err: None,
            })
    }
}

// Это обертка над хешкартой, нужна чтобы централизовать обработку ошибок.
// В противном случае каждая попытка прочитать данные из Hmap потребует индивидуальный unwrap с конвертацией в ошибку
#[derive(Debug)]
pub struct TagAddressMap {
    data: HashMap<TagID, (usize, usize)>,
}

impl TagAddressMap {
    pub fn new() -> Self {
        Self {
            data: HashMap::new(),
        }
    }
    pub fn get(&self, key: &TagID) -> Result<&(usize, usize), Error<'static>> {
        self.data.get(key).ok_or_else(|| Error::InternalLogic {
            tech_descr: format!(r#"Хешкарта не содержит ключ "{}""#, key.as_str()),
            err: None,
        })
    }
    pub fn insert(&mut self, key: TagID, data: (usize, usize)) {
        self.data.insert(key, data);
    }
}
