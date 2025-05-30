use crate::atoms;
use crate::UmyaSpreadsheet;
use rustler::{Atom, Error as NifError, NifResult, ResourceArc};

/// Helper function to find sheet index by name
fn find_sheet_index_by_name(
    guard: &std::sync::MutexGuard<umya_spreadsheet::Spreadsheet>,
    sheet_name: &str,
) -> Option<usize> {
    guard
        .get_sheet_collection_no_check()
        .iter()
        .position(|sheet| sheet.get_name() == sheet_name)
}

/// Helper function to ensure a worksheet is deserialized
fn ensure_worksheet_deserialized(
    guard: &mut std::sync::MutexGuard<umya_spreadsheet::Spreadsheet>,
    sheet_index: &usize,
) {
    // Call read_sheet to deserialize the worksheet
    // This is safe to call even if already deserialized
    guard.read_sheet(*sheet_index);
}

/// Get the value of a cell
#[rustler::nif]
pub fn get_cell_value(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_reference: String,
) -> NifResult<String> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    // Find sheet index by name
    let sheet_index = match find_sheet_index_by_name(&guard, &sheet_name) {
        Some(index) => index,
        None => {
            return Err(NifError::Term(Box::new((
                atoms::error(),
                "Sheet not found".to_string(),
            ))))
        }
    };

    // Ensure the worksheet is deserialized before accessing it
    ensure_worksheet_deserialized(&mut guard, &sheet_index);

    // Use get_sheet_mut which automatically deserializes the worksheet
    if let Some(sheet) = guard.get_sheet_mut(&sheet_index) {
        match sheet.get_cell(cell_reference.as_str()) {
            Some(cell) => Ok(cell.get_value().to_string()),
            None => Ok("".to_string()), // Empty string for non-existent cells
        }
    } else {
        Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        ))))
    }
}

/// Set the value of a cell
#[rustler::nif]
pub fn set_cell_value(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_reference: String,
    value: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    // Find sheet index by name
    let sheet_index = match find_sheet_index_by_name(&guard, &sheet_name) {
        Some(index) => index,
        None => {
            return Err(NifError::Term(Box::new((
                atoms::error(),
                "Sheet not found".to_string(),
            ))))
        }
    };

    // Ensure the worksheet is deserialized before accessing it
    ensure_worksheet_deserialized(&mut guard, &sheet_index);

    // Use get_sheet_mut which automatically deserializes the worksheet
    if let Some(sheet) = guard.get_sheet_mut(&sheet_index) {
        // Create a new or get existing cell
        let cell = sheet.get_cell_mut(cell_reference.as_str());

        // Set the value based on the content (numbers, booleans, etc.)
        if value.trim().to_lowercase() == "true" {
            cell.set_value("TRUE");
        } else if value.trim().to_lowercase() == "false" {
            cell.set_value("FALSE");
        } else if let Ok(int_val) = value.parse::<i64>() {
            cell.set_value(&(int_val as f64).to_string());
        } else if let Ok(float_val) = value.parse::<f64>() {
            cell.set_value(&float_val.to_string());
        } else {
            cell.set_value(&value);
        }

        Ok(atoms::ok())
    } else {
        Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        ))))
    }
}

/// Remove a cell from a sheet
#[rustler::nif]
pub fn remove_cell(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_reference: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    // Find sheet index by name
    let sheet_index = match find_sheet_index_by_name(&guard, &sheet_name) {
        Some(index) => index,
        None => {
            return Err(NifError::Term(Box::new((
                atoms::error(),
                "Sheet not found".to_string(),
            ))))
        }
    };

    // Ensure the worksheet is deserialized before accessing it
    ensure_worksheet_deserialized(&mut guard, &sheet_index);

    // Use get_sheet_mut which automatically deserializes the worksheet
    if let Some(sheet) = guard.get_sheet_mut(&sheet_index) {
        sheet.remove_cell(cell_reference.as_str());
        Ok(atoms::ok())
    } else {
        Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        ))))
    }
}

/// Get the formatted value of a cell (applies number format)
#[rustler::nif]
pub fn get_formatted_value(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_reference: String,
) -> NifResult<String> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    // Find sheet index by name
    let sheet_index = match find_sheet_index_by_name(&guard, &sheet_name) {
        Some(index) => index,
        None => {
            return Err(NifError::Term(Box::new((
                atoms::error(),
                "Sheet not found".to_string(),
            ))))
        }
    };

    // Ensure the worksheet is deserialized before accessing it
    ensure_worksheet_deserialized(&mut guard, &sheet_index);

    // Use get_sheet_mut which automatically deserializes the worksheet
    if let Some(sheet) = guard.get_sheet_mut(&sheet_index) {
        match sheet.get_cell(cell_reference.as_str()) {
            Some(_cell) => {
                // Get the formatted value directly
                Ok(sheet.get_formatted_value(cell_reference.as_str()))
            }
            None => Ok("".to_string()), // Empty string for non-existent cells
        }
    } else {
        Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        ))))
    }
}

/// Set the number format for a cell
#[rustler::nif]
pub fn set_number_format(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_reference: String,
    format_code: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    // Find sheet index by name
    let sheet_index = match find_sheet_index_by_name(&guard, &sheet_name) {
        Some(index) => index,
        None => {
            return Err(NifError::Term(Box::new((
                atoms::error(),
                "Sheet not found".to_string(),
            ))))
        }
    };

    // Ensure the worksheet is deserialized before accessing it
    ensure_worksheet_deserialized(&mut guard, &sheet_index);

    // Use get_sheet_mut which automatically deserializes the worksheet
    if let Some(sheet) = guard.get_sheet_mut(&sheet_index) {
        let cell = sheet.get_cell_mut(cell_reference.as_str());

        // Clone the existing style to modify it
        let mut style = cell.get_style().clone();
        style.get_number_format_mut().set_format_code(&format_code);
        cell.set_style(style);

        Ok(atoms::ok())
    } else {
        Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        ))))
    }
}
