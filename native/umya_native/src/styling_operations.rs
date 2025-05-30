use rustler::{Atom, Error as NifError, NifResult, ResourceArc};

use crate::atoms;
use crate::UmyaSpreadsheet;
use crate::helpers::style_helpers;

/// Set the font color for a cell
#[rustler::nif]
pub fn set_font_color(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_reference: String,
    color: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let cell = sheet.get_cell_mut(cell_reference.as_str());

            // Parse color and apply it
            let color_obj = style_helpers::parse_color(&color);
            let mut style = cell.get_style().clone();
            let mut font = style.get_font().map(|f| f.clone()).unwrap_or_default();
            font.set_color(color_obj);
            style.set_font(font);
            cell.set_style(style);

            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Set the font size for a cell
#[rustler::nif]
pub fn set_font_size(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_reference: String,
    font_size: f64,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let cell = sheet.get_cell_mut(cell_reference.as_str());

            let mut style = cell.get_style().clone();
            let mut font = style.get_font().map(|f| f.clone()).unwrap_or_default();
            font.set_size(font_size);
            style.set_font(font);
            cell.set_style(style);

            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Set the font bold status for a cell
#[rustler::nif]
pub fn set_font_bold(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_reference: String,
    is_bold: bool,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let cell = sheet.get_cell_mut(cell_reference.as_str());

            let mut style = cell.get_style().clone();
            let mut font = style.get_font().map(|f| f.clone()).unwrap_or_default();
            font.set_bold(is_bold);
            style.set_font(font);
            cell.set_style(style);

            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Set the font name for a cell
#[rustler::nif]
pub fn set_font_name(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_reference: String,
    font_name: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let cell = sheet.get_cell_mut(cell_reference.as_str());

            let mut style = cell.get_style().clone();
            let mut font = style.get_font().map(|f| f.clone()).unwrap_or_default();
            font.set_name(&font_name);
            style.set_font(font);
            cell.set_style(style);

            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Copy styling from one row to another
#[rustler::nif]
pub fn copy_row_styling(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    source_row: u32,
    target_row: u32,
    start_column: Option<u32>,
    end_column: Option<u32>,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Use the worksheet's copy_row_styling method
            sheet.copy_row_styling(
                &source_row,
                &target_row,
                start_column.as_ref(),
                end_column.as_ref(),
            );

            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Copy styling from one column to another
#[rustler::nif]
pub fn copy_column_styling(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    source_column: u32,
    target_column: u32,
    start_row: Option<u32>,
    end_row: Option<u32>,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Use the worksheet's copy_col_styling method
            sheet.copy_col_styling(
                &source_column,
                &target_column,
                start_row.as_ref(),
                end_row.as_ref(),
            );

            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}
