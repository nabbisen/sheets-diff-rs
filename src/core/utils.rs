/// filter sheets whose name is equal
pub fn filter_same_name_sheets<'a>(
    old_sheets: &'a Vec<String>,
    new_sheets: &'a Vec<String>,
) -> Vec<String> {
    old_sheets
        .iter()
        .filter(|s| new_sheets.contains(s))
        .map(|s| s.to_owned())
        .collect()
}

/// get range to compare
/// return: (start_row, start_col, end_row, end_col)
pub fn diff_range<'a>(
    old_start: Option<(u32, u32)>,
    new_start: Option<(u32, u32)>,
    old_end: Option<(u32, u32)>,
    new_end: Option<(u32, u32)>,
) -> (u32, u32, u32, u32) {
    let (old_start_row, old_start_col) = match old_start {
        Some((row, col)) => (row, col),
        None => (u32::MAX, u32::MAX),
    };
    let (new_start_row, new_start_col) = match new_start {
        Some((row, col)) => (row, col),
        None => (u32::MAX, u32::MAX),
    };
    let (old_end_row, old_end_col) = match old_end {
        Some((row, col)) => (row, col),
        None => (u32::MIN, u32::MIN),
    };
    let (new_end_row, new_end_col) = match new_end {
        Some((row, col)) => (row, col),
        None => (u32::MIN, u32::MIN),
    };
    let start_row = old_start_row.min(new_start_row);
    let start_col = old_start_col.min(new_start_col);
    let end_row = old_end_row.max(new_end_row);
    let end_col = old_end_col.max(new_end_col);

    (start_row, start_col, end_row + 1, end_col + 1)
}

/// convert (row, col) to cell address str
pub fn cell_pos_to_address(row: usize, col: usize) -> String {
    let col_letter = (col as u8 - 1) / 26;
    let col_index = (col as u8 - 1) % 26;

    let col_char = if col_letter == 0 {
        ((b'A' + col_index) as char).to_string()
    } else {
        let first_char = (b'A' + col_letter - 1) as char;
        let second_char = (b'A' + col_index) as char;
        format!("{}{}", first_char, second_char)
    };

    format!("{}{}", col_char, row)
}
