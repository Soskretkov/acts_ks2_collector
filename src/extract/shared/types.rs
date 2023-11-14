use crate::errors::Error;
use std::collections::HashMap;

// это обертка над хешкартой нужна чтобы централизовать обработку ошибок
// в противном случае каждая попытка прочитать данные требует свой unwrap() с ловлей ошибки
pub struct SearchPoint {
    data: HashMap<TagID, (usize, usize)>,
}

impl SearchPoint {
    pub fn new() -> Self {
        Self {
            data: HashMap::new(),
        }
    }
    pub fn get(&self, key: &TagID) -> Result<&(usize, usize), Error<'static>> {
        self.data.get(key).ok_or_else(|| Error::InternalLogic {
            tech_descr: format!(r#"Хешкарта не содержит ключа "{}""#, key.as_str()),
            err: None,
        })
    }
    pub fn insert(&mut self, key: TagID, data: (usize, usize)) {
        self.data.insert(key, data);
    }
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
