pub fn get_xl_column_letter(zero_based_column: u16) -> String {
    let integer = zero_based_column / 26;
    let remainder = (zero_based_column % 26) as u8;
    let ch = char::from(remainder + 65).to_ascii_uppercase().to_string();

    if integer == 0 {
        return ch;
    }

    get_xl_column_letter(integer - 1) + &ch
}


#[cfg(test)]
mod tests {
    #[test]
    fn column_in_excel_with_letters_01() {
        use super::get_xl_column_letter;
        let result = get_xl_column_letter(886);
        assert_eq!(result, "AHC".to_string());
    }
    #[test]
    fn column_in_excel_with_letters_02() {
        use super::get_xl_column_letter;
        let result = get_xl_column_letter(1465);
        assert_eq!(result, "BDJ".to_string());
    }
}