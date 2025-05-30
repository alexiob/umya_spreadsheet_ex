use rustler::{Atom, Error as NifError, NifResult, ResourceArc};

use crate::atoms;
use crate::UmyaSpreadsheet;

/// Set the height of a row
#[rustler::nif]
pub fn set_row_height(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    row_number: u32,
    height: f64,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // get_row_dimension_mut returns &mut Row directly, not Option
            let row_dim = sheet.get_row_dimension_mut(&row_number);
            row_dim.set_height(height);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new("Sheet not found".to_string()))),
    }
}

/// Set the style for an entire row
#[rustler::nif]
pub fn set_row_style(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    row_number: u32,
    bg_color: String,
    font_color: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let row_dim = sheet.get_row_dimension_mut(&row_number);

            // Create a new style with the background and font colors
            let mut style = umya_spreadsheet::Style::default();

            // Set background color if provided and not empty
            if !bg_color.is_empty() {
                let bg_color_obj = crate::helpers::style_helpers::parse_color(&bg_color);
                let mut fill = style.get_fill().cloned().unwrap_or_default();
                let mut pattern_fill = fill.get_pattern_fill().cloned().unwrap_or_default();
                pattern_fill.set_foreground_color(bg_color_obj);
                pattern_fill.set_pattern_type(umya_spreadsheet::PatternValues::Solid);
                fill.set_pattern_fill(pattern_fill);
                style.set_fill(fill);
            }

            // Set font color if provided and not empty
            if !font_color.is_empty() {
                let font_color_obj = crate::helpers::style_helpers::parse_color(&font_color);
                let mut font = style.get_font().cloned().unwrap_or_default();
                font.set_color(font_color_obj);
                style.set_font(font);
            }

            row_dim.set_style(style);

            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new("Sheet not found".to_string()))),
    }
}

/// Set the width of a column
#[rustler::nif]
pub fn set_column_width(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    column: String,
    width: f64,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // get_column_dimension_mut returns &mut Column directly, not Option
            let col_dim = sheet.get_column_dimension_mut(&column);
            col_dim.set_width(width);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new("Sheet not found".to_string()))),
    }
}

/// Set auto width for a column
#[rustler::nif]
pub fn set_column_auto_width(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    column: String,
    is_auto_width: bool,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // get_column_dimension_mut returns &mut Column directly, not Option
            let column_dimension = sheet.get_column_dimension_mut(&column);
            column_dimension.set_best_fit(is_auto_width);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new("Sheet not found".to_string()))),
    }
}

/// Get the width of a column
#[rustler::nif]
pub fn get_column_width(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    column: String,
) -> NifResult<f64> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            match sheet.get_column_dimension(&column) {
                Some(column_dim) => {
                    let width = column_dim.get_width();
                    Ok(*width)
                }
                None => Ok(8.43), // Default column width
            }
        }
        None => Err(NifError::Term(Box::new("Sheet not found".to_string()))),
    }
}

/// Get column auto width setting
#[rustler::nif]
pub fn get_column_auto_width(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    column: String,
) -> NifResult<bool> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            match sheet.get_column_dimension(&column) {
                Some(column_dim) => {
                    // Return best_fit value since that's what gets persisted
                    let best_fit = column_dim.get_best_fit();
                    Ok(*best_fit)
                }
                None => Ok(false), // Default best_fit is false
            }
        }
        None => Err(NifError::Term(Box::new("Sheet not found".to_string()))),
    }
}

/// Get if a column is hidden
#[rustler::nif]
pub fn get_column_hidden(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    column: String,
) -> NifResult<bool> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            match sheet.get_column_dimension(&column) {
                Some(column_dim) => {
                    let hidden = column_dim.get_hidden();
                    Ok(*hidden)
                }
                None => Ok(false), // Default hidden is false
            }
        }
        None => Err(NifError::Term(Box::new("Sheet not found".to_string()))),
    }
}

/// Get the height of a row
#[rustler::nif]
pub fn get_row_height(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    row_number: u32,
) -> NifResult<f64> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            match sheet.get_row_dimension(&row_number) {
                Some(row_dim) => {
                    let height = row_dim.get_height();
                    Ok(*height)
                }
                None => Ok(15.0), // Default row height
            }
        }
        None => Err(NifError::Term(Box::new("Sheet not found".to_string()))),
    }
}

/// Get if a row is hidden
#[rustler::nif]
pub fn get_row_hidden(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    row_number: u32,
) -> NifResult<bool> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            match sheet.get_row_dimension(&row_number) {
                Some(row_dim) => {
                    let hidden = row_dim.get_hidden();
                    Ok(*hidden)
                }
                None => Ok(false), // Default hidden is false
            }
        }
        None => Err(NifError::Term(Box::new("Sheet not found".to_string()))),
    }
}

/// Remove rows from a worksheet
#[rustler::nif]
pub fn remove_row(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    row_index: u32,
    amount: u32,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            sheet.remove_row(&row_index, &amount);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Remove columns from a worksheet
#[rustler::nif]
pub fn remove_column(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    column_letter: String,
    amount: u32,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            sheet.remove_column(column_letter.as_str(), &amount);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Remove columns from a worksheet by index
#[rustler::nif]
pub fn remove_column_by_index(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    column_index: u32,
    amount: u32,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            sheet.remove_column_by_index(&column_index, &amount);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}
