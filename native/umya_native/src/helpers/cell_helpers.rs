use umya_spreadsheet::helper::coordinate::index_from_coordinate;

/// Check if a cell address is within a given range
///
/// # Arguments
/// * `cell_address` - The cell address to check (e.g., "A1")
/// * `range` - The range to check against (e.g., "A1:C3")
///
/// # Returns
/// `true` if the cell is within the range, `false` otherwise
pub fn check_inside(cell_address: &str, range: &str) -> bool {
    // Parse the cell address
    let (cell_col, cell_row, ..) = index_from_coordinate(cell_address);

    // Parse the range
    let (row_start, row_end, col_start, col_end) =
        umya_spreadsheet::helper::range::get_start_and_end_point(range);

    // Both cell column and row must be defined
    if cell_col.is_none() || cell_row.is_none() {
        return false;
    }

    // Check if the cell is within the range
    let col = cell_col.unwrap();
    let row = cell_row.unwrap();

    col >= col_start && col <= col_end && row >= row_start && row <= row_end
}

/// Check if a cell address is within a given range
///
/// # Arguments
/// * `cell_address` - The cell address to check (e.g., "A1")
/// * `range` - The range to check against (e.g., "A1:C3")
///
/// # Returns
/// `true` if the cell is within the range, `false` otherwise
pub fn is_cell_in_range(cell_address: &str, range: &str) -> bool {
    check_inside(cell_address, range)
}
