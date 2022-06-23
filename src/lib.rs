use std::error::Error;
#[derive(Debug)]
// pub enum ErrKind {
//     Handled,
//     Fatal,
// }

pub enum ErrName {
    shifted_columns_in_header,
    sheet_not_contain_all_necessary_data,
    calamine_sheet_of_the_book_is_undetectable,
    calamine_sheet_of_the_book_is_unreadable(calamine::XlsxError),
    calamine,
}

#[derive(Debug)]
pub struct ErrDescription {
    pub name: ErrName,
    pub description: Option<String>,
}
