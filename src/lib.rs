#[derive(Debug)]
// pub enum ErrKind {
//     Handled,
//     Fatal,
// }

pub enum ErrName {
    ShiftedColumnsInHeader,
    SheetNotContainAllNecessaryData,
    CalamineSheetOfTheBookIsUndetectable,
    CalamineSheetOfTheBookIsUnreadable(calamine::XlsxError),
    Calamine,
}

#[derive(Debug)]
pub struct ErrDescription {
    pub name: ErrName,
    pub description: Option<String>,
}

pub fn variant_eq<T>(first: &T, second: &T) -> bool {
    std::mem::discriminant(first) == std::mem::discriminant(second)
}
