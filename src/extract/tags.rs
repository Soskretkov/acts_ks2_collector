use crate::errors::Error;
use std::collections::HashMap;

// перечислены в порядке вхождения слева на право и сверху вниз на листе Excel (вход по строкам важен для валидации)
// группировка по строке и столбцу для валидации в будующих версиях программы (не реализовано)
#[rustfmt::skip]
pub const TAG_INFO_ARRAY: [TagInfo; 10] = [
    TagInfo { is_required: false, group_by_row: None,                   group_by_col: Some(Column::Initial),  id: TagID::Исполнитель },
    TagInfo { is_required: true,  group_by_row: None,                   group_by_col: Some(Column::Initial),  id: TagID::Стройка },
    TagInfo { is_required: true,  group_by_row: None,                   group_by_col: Some(Column::Initial),  id: TagID::Объект },
    TagInfo { is_required: true,  group_by_row: None,                   group_by_col: Some(Column::Contract), id: TagID::ДоговорПодряда },
    TagInfo { is_required: true,  group_by_row: None,                   group_by_col: Some(Column::Contract), id: TagID::ДопСоглашение },
    TagInfo { is_required: true,  group_by_row: None,                   group_by_col: None,                   id: TagID::НомерДокумента },
    TagInfo { is_required: true,  group_by_row: Some(Row::TableHeader), group_by_col: None,                   id: TagID::НаименованиеРаботИЗатрат },
    TagInfo { is_required: false, group_by_row: Some(Row::TableHeader), group_by_col: None,                   id: TagID::ЗтрВсего },
    TagInfo { is_required: false, group_by_row: None,                   group_by_col: Some(Column::Initial),  id: TagID::ИтогоПоАкту },
    TagInfo { is_required: true,  group_by_row: None,                   group_by_col: None,                   id: TagID::СтоимостьМатериальныхРесурсовВсего },
];

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
    ЗтрВсего,
    ИтогоПоАкту,
    СтоимостьМатериальныхРесурсовВсего,
}

#[rustfmt::skip]
impl TagID {
    pub fn as_str(&self) -> &'static str {
        match self {
            TagID::Исполнитель => "исполнитель",
            TagID::Стройка => "стройка",
            TagID::Объект => "объект",
            TagID::ДоговорПодряда => "договор подряда",
            TagID::ДопСоглашение => "доп. соглашение",
            TagID::НомерДокумента => "номер документа",
            TagID::НаименованиеРаботИЗатрат => "наименование работ и затрат",
            TagID::ЗтрВсего => "зтр всего чел.-час",
            TagID::ИтогоПоАкту => "итого по акту:",
            TagID::СтоимостьМатериальныхРесурсовВсего => "стоимость материальных ресурсов (всего)",
        }
    }
}

#[derive(Clone, Copy)]
pub struct TagInfo {
    pub id: TagID,
    pub is_required: bool,
    pub group_by_row: Option<Row>,
    pub group_by_col: Option<Column>,
}

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

// это обертка над хешкартой нужна чтобы централизовать обработку ошибок
// в противном случае каждая попытка прочитать данные требует свой unwrap с ловлей ошибки
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
