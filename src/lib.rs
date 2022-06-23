
use std::error::Error;
pub enum ErrKind {
    Handled,
    Fatal,
}

pub enum ErrName {
    shifted_columns_in_header,
    sheet_not_contain_all_necessary_data,
    book_not_contain_requested_sheet,
    calamine,


}

pub struct ErrDescription {
    name: ErrName,
    kind: ErrKind,
    description: Option<String>,
    err: Option<Box<dyn Error>>,
}