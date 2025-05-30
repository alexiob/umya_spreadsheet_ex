use rustler::{Atom, Error as NifError, NifResult, ResourceArc};
use umya_spreadsheet;

use crate::atoms;
use crate::UmyaSpreadsheet;

/// Helper function to ensure a worksheet is deserialized
fn ensure_worksheet_deserialized(
    guard: &mut std::sync::MutexGuard<umya_spreadsheet::Spreadsheet>,
    sheet_index: &usize,
) {
    // Call read_sheet to deserialize the worksheet
    // This is safe to call even if already deserialized
    guard.read_sheet(*sheet_index);
}

/// Get all sheet names from a spreadsheet
#[rustler::nif]
pub fn get_sheet_names(resource: ResourceArc<UmyaSpreadsheet>) -> NifResult<Vec<String>> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    // Safely get sheet names by examining the workbook structure directly
    let mut sheet_names = Vec::new();
    let count = guard.get_sheet_count();

    for i in 0..count {
        // Always ensure the worksheet is deserialized first
        ensure_worksheet_deserialized(&mut guard, &i);

        // Try to get the sheet name using a safer approach
        // First try the normal path, but catch panics
        let name = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            if let Some(sheet) = guard.get_sheet(&i) {
                sheet.get_name().to_string()
            } else {
                format!("Sheet{}", i + 1) // Fallback name
            }
        }));

        match name {
            Ok(sheet_name) => sheet_names.push(sheet_name),
            Err(_) => {
                // If there's a panic, use a default name
                sheet_names.push(format!("Sheet{}", i + 1));
            }
        }
    }

    Ok(sheet_names)
}

/// Get the number of sheets in a spreadsheet
#[rustler::nif]
pub fn get_sheet_count(resource: ResourceArc<UmyaSpreadsheet>) -> NifResult<usize> {
    let guard = resource.spreadsheet.lock().unwrap();
    Ok(guard.get_sheet_count())
}

/// Get the state of a sheet (visible, hidden, very hidden)
#[rustler::nif]
pub fn get_sheet_state(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> NifResult<(Atom, String)> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let state = sheet.get_sheet_state();
            // get_sheet_state() returns &str directly now
            // If state is empty, default to "visible"
            let state_str = if state.is_empty() { "visible" } else { state };
            Ok((atoms::ok(), state_str.to_string()))
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Get sheet protection status
#[rustler::nif]
pub fn get_sheet_protection(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> NifResult<(Atom, std::collections::HashMap<String, bool>)> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let mut protection_map = std::collections::HashMap::new();

            if let Some(protection) = sheet.get_sheet_protection() {
                protection_map.insert("sheet".to_string(), *protection.get_sheet());
                protection_map.insert("objects".to_string(), *protection.get_objects());
                protection_map.insert("scenarios".to_string(), *protection.get_scenarios());
                protection_map.insert("format_cells".to_string(), *protection.get_format_cells());
                protection_map.insert("format_columns".to_string(), *protection.get_format_columns());
                protection_map.insert("format_rows".to_string(), *protection.get_format_rows());
                protection_map.insert("insert_columns".to_string(), *protection.get_insert_columns());
                protection_map.insert("insert_rows".to_string(), *protection.get_insert_rows());
                protection_map.insert("insert_hyperlinks".to_string(), *protection.get_insert_hyperlinks());
                protection_map.insert("delete_columns".to_string(), *protection.get_delete_columns());
                protection_map.insert("delete_rows".to_string(), *protection.get_delete_rows());
                protection_map.insert("select_locked_cells".to_string(), *protection.get_select_locked_cells());
                protection_map.insert("sort".to_string(), *protection.get_sort());
                protection_map.insert("auto_filter".to_string(), *protection.get_auto_filter());
                protection_map.insert("pivot_tables".to_string(), *protection.get_pivot_tables());
                protection_map.insert("select_unlocked_cells".to_string(), *protection.get_select_unlocked_cells());
            } else {
                // Return default protection settings if none exist
                protection_map.insert("sheet".to_string(), false);
                protection_map.insert("objects".to_string(), false);
                protection_map.insert("scenarios".to_string(), false);
                protection_map.insert("format_cells".to_string(), true);
                protection_map.insert("format_columns".to_string(), true);
                protection_map.insert("format_rows".to_string(), true);
                protection_map.insert("insert_columns".to_string(), true);
                protection_map.insert("insert_rows".to_string(), true);
                protection_map.insert("insert_hyperlinks".to_string(), true);
                protection_map.insert("delete_columns".to_string(), true);
                protection_map.insert("delete_rows".to_string(), true);
                protection_map.insert("select_locked_cells".to_string(), true);
                protection_map.insert("sort".to_string(), true);
                protection_map.insert("auto_filter".to_string(), true);
                protection_map.insert("pivot_tables".to_string(), true);
                protection_map.insert("select_unlocked_cells".to_string(), true);
            }

            Ok((atoms::ok(), protection_map))
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Get all merged cells in a sheet
#[rustler::nif]
pub fn get_merge_cells(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> NifResult<Vec<String>> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let merge_cells = sheet.get_merge_cells();
            let ranges: Vec<String> = merge_cells
                .iter()
                .map(|mc| mc.get_range())
                .collect();
            Ok(ranges)
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Add a new sheet to the spreadsheet
#[rustler::nif]
pub fn add_sheet(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    // Check if a sheet with this name already exists
    if guard.get_sheet_by_name_mut(&sheet_name).is_some() {
        return Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet already exists".to_string(),
        ))));
    }

    // Create a new worksheet and add it to the spreadsheet
    let mut new_sheet = umya_spreadsheet::Worksheet::default();
    new_sheet.set_name(&sheet_name);
    let _ = guard.add_sheet(new_sheet);
    Ok(atoms::ok())
}

/// Rename an existing sheet
#[rustler::nif]
pub fn rename_sheet(
    resource: ResourceArc<UmyaSpreadsheet>,
    old_name: String,
    new_name: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    // Check if source sheet exists
    if guard.get_sheet_by_name_mut(&old_name).is_none() {
        return Err(NifError::Term(Box::new((
            atoms::error(),
            "Source sheet not found".to_string(),
        ))));
    }

    // Check if target name already exists
    if guard.get_sheet_by_name_mut(&new_name).is_some() {
        return Err(NifError::Term(Box::new((
            atoms::error(),
            "Target sheet already exists".to_string(),
        ))));
    }

    // Rename sheet
    if let Some(sheet) = guard.get_sheet_by_name_mut(&old_name) {
        sheet.set_name(&new_name);
    }
    Ok(atoms::ok())
}

/// Move a range of cells to a new location
#[rustler::nif]
pub fn move_range(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    range: String,
    row: i32,
    column: i32,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Use the built-in move_range function from umya-spreadsheet
            sheet.move_range(&range, &row, &column);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Clone an existing sheet with a new name
#[rustler::nif]
pub fn clone_sheet(
    resource: ResourceArc<UmyaSpreadsheet>,
    source_sheet_name: String,
    new_sheet_name: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&source_sheet_name) {
        Some(source_sheet) => {
            let mut cloned_sheet = source_sheet.clone();
            cloned_sheet.set_name(&new_sheet_name);
            let _ = guard.add_sheet(cloned_sheet);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Source sheet not found".to_string(),
        )))),
    }
}

/// Remove a sheet from the spreadsheet
#[rustler::nif]
pub fn remove_sheet(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.remove_sheet_by_name(&sheet_name) {
        Ok(_) => Ok(atoms::ok()),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found or cannot be removed".to_string(),
        )))),
    }
}

/// Insert new rows into a sheet
#[rustler::nif]
pub fn insert_new_row(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    row_index: u32,
    amount: u32,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            sheet.insert_new_row(&row_index, &amount);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Insert new columns into a sheet
#[rustler::nif]
pub fn insert_new_column(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    column_letter: String,
    amount: u32,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Use the column letter directly
            sheet.insert_new_column(column_letter.as_str(), &amount);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Set sheet protection
#[rustler::nif]
pub fn set_sheet_protection(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    password: Option<String>,
    protected: bool,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let protection = sheet.get_sheet_protection_mut();
            protection.set_sheet(protected);

            if let Some(pwd) = password {
                protection.set_password(&pwd);
            }

            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Add merge cells to a sheet
#[rustler::nif]
pub fn add_merge_cells(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    range: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Create a new Range and add it to merge cells
            let mut new_range = umya_spreadsheet::Range::default();
            new_range.set_range(range.as_str());
            sheet.get_merge_cells_mut().push(new_range);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Set sheet state (visible, hidden, very hidden)
#[rustler::nif]
pub fn set_sheet_state(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    state: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            sheet.set_sheet_state(state);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}
