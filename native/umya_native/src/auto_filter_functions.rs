use crate::atoms;
use crate::UmyaSpreadsheet;
use rustler::{Atom, Error as NifError, NifResult};
use std::panic::{self, AssertUnwindSafe};

#[rustler::nif]
pub fn set_auto_filter(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    range: String,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<(), String> {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        // Validate inputs
        if range.trim().is_empty() {
            return Err("Range cannot be empty".to_string());
        }

        // Get sheet by name
        let sheet = spreadsheet.get_sheet_by_name_mut(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        // Set auto filter for the specified range
        sheet.set_auto_filter(range);
        Ok(())
    }));

    match result {
        Ok(Ok(())) => Ok(atoms::ok()),
        Ok(Err(err_msg)) => Err(NifError::Term(Box::new((atoms::error(), err_msg)))),
        Err(_) => Err(NifError::Term(Box::new((atoms::error(), "Error occurred in set_auto_filter operation".to_string())))),
    }
}

#[rustler::nif]
pub fn remove_auto_filter(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<(), String> {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        // Get sheet by name
        let sheet = spreadsheet.get_sheet_by_name_mut(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        // Remove auto filter
        sheet.remove_auto_filter();
        Ok(())
    }));

    match result {
        Ok(Ok(())) => Ok(atoms::ok()),
        Ok(Err(err_msg)) => Err(NifError::Term(Box::new((atoms::error(), err_msg)))),
        Err(_) => Err(NifError::Term(Box::new((atoms::error(), "Error occurred in remove_auto_filter operation".to_string())))),
    }
}

#[rustler::nif]
pub fn has_auto_filter(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> bool {
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
        Ok(Ok(has_filter)) => has_filter,
        Ok(Err(_msg)) => {
            // Silent error handling - return default value
            false
        }
        Err(_) => {
            // Silent error handling - return default value
            false
        }
    }
}

#[rustler::nif]
pub fn get_auto_filter_range(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> Option<String> {
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
        Ok(Ok(range_opt)) => range_opt,
        Ok(Err(_msg)) => {
            // Silent error handling - return default value
            None
        }
        Err(_) => {
            // Silent error handling - return default value
            None
        }
    }
}
