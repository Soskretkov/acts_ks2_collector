use std::error::Error;
#[derive(Debug)]
// pub enum ErrKind {
//     Handled,
//     Fatal,
// }

pub enum ErrName {
    Shifted_columns_in_header,
    Sheet_not_contain_all_necessary_data,
    Calamine_sheet_of_the_book_is_undetectable,
    Calamine_sheet_of_the_book_is_unreadable(calamine::XlsxError),
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