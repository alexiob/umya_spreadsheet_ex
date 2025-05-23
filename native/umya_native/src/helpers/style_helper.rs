use umya_spreadsheet::{Color, PatternFill};

/// Creates a new pattern fill using the specified color
///
/// # Arguments
///
/// * `color` - The Color object to use for the pattern fill
///
/// # Returns
///
/// A PatternFill instance with the specified color as foreground color
pub fn create_pattern_fill(color: Color) -> PatternFill {
    PatternFill::default()
        .set_foreground_color(color)
        .to_owned()
}

/// Sets the style properties on a cell
///
/// # Arguments
///
/// * `sheet` - Mutable reference to a worksheet
/// * `cell_address` - Cell address in A1 notation
/// * `bg_color` - Optional background color as Color object
/// * `font_color` - Optional font color as Color object
/// * `font_size` - Optional font size as f64
/// * `is_bold` - Optional boolean to set bold text
///
pub fn apply_cell_style(
    sheet: &mut umya_spreadsheet::Worksheet,
    cell_address: &str,
    bg_color: Option<Color>,
    font_color: Option<Color>,
    font_size: Option<f64>,
    is_bold: Option<bool>,
) {
    let style = sheet.get_style_mut(cell_address);

    // Set background color if provided
    if let Some(color) = bg_color {
        let fill = create_pattern_fill(color);
        style.get_fill_mut().set_pattern_fill(fill);
    }

    // Set font color if provided
    if let Some(color) = font_color {
        style.get_font_mut().set_color(color);
    }

    // Set font size if provided
    if let Some(size) = font_size {
        style.get_font_mut().set_size(size);
    }

    // Set bold if provided
    if let Some(bold) = is_bold {
        style.get_font_mut().set_bold(bold);
    }
}

/// Applies a style to an entire row
///
/// # Arguments
///
/// * `sheet` - Mutable reference to a worksheet
/// * `row_number` - The row number (1-based)
/// * `bg_color` - Optional background color as Color object
/// * `font_color` - Optional font color as Color object
/// * `font_size` - Optional font size as f64
/// * `is_bold` - Optional boolean to set bold text
///
pub fn apply_row_style(
    sheet: &mut umya_spreadsheet::Worksheet,
    row_number: u32,
    bg_color: Option<Color>,
    font_color: Option<Color>,
    font_size: Option<f64>,
    is_bold: Option<bool>,
) {
    let row_dim = sheet.get_row_dimension_mut(&row_number);
    let style = row_dim.get_style_mut();

    // Set background color if provided
    if let Some(color) = bg_color {
        let fill = create_pattern_fill(color);
        style.get_fill_mut().set_pattern_fill(fill);
    }

    // Set font color if provided
    if let Some(color) = font_color {
        style.get_font_mut().set_color(color);
    }

    // Set font size if provided
    if let Some(size) = font_size {
        style.get_font_mut().set_size(size);
    }

    // Set bold if provided
    if let Some(bold) = is_bold {
        style.get_font_mut().set_bold(bold);
    }
}
