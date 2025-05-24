use crate::atoms;
use crate::helpers::error_helper::handle_error;
use crate::UmyaSpreadsheet;
use rustler::{Atom, NifResult};
use std::panic::{self, AssertUnwindSafe};

#[rustler::nif]
pub fn set_auto_filter(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    range: String,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) {
            // Set auto filter for the specified range
            sheet.set_auto_filter(range);
            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => handle_error(&msg),
        Err(_) => handle_error("Panic occurred in set_auto_filter"),
    }
}

#[rustler::nif]
pub fn remove_auto_filter(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) {
            // Remove auto filter
            sheet.remove_auto_filter();
            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => handle_error(&msg),
        Err(_) => handle_error("Panic occurred in remove_auto_filter"),
    }
}

#[rustler::nif]
pub fn has_auto_filter(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> NifResult<bool> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name(&sheet_name) {
            // Check if auto filter exists
            Ok(sheet.get_auto_filter().is_some())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(has_filter)) => Ok(has_filter),
        Ok(Err(msg)) => {
            let _: Atom = handle_error(&msg)?;
            Ok(false)
        }
        Err(_) => {
            let _: Atom = handle_error("Panic occurred in has_auto_filter")?;
            Ok(false)
        }
    }
}

#[rustler::nif]
pub fn get_auto_filter_range(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> NifResult<Option<String>> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name(&sheet_name) {
            // Get auto filter range if it exists
            if let Some(auto_filter) = sheet.get_auto_filter() {
                Ok(Some(auto_filter.get_range().get_range().to_string()))
            } else {
                Ok(None)
            }
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(range_opt)) => Ok(range_opt),
        Ok(Err(msg)) => {
            let _: Atom = handle_error(&msg)?;
            Ok(None)
        }
        Err(_) => {
            let _: Atom = handle_error("Panic occurred in get_auto_filter_range")?;
            Ok(None)
        }
    }
}
