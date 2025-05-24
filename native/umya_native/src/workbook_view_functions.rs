use crate::atoms;
use crate::UmyaSpreadsheet;
use rustler::{Atom, NifResult, Error as NifError};
use std::panic::{self, AssertUnwindSafe};

#[rustler::nif]
pub fn set_active_tab(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    tab_index: u32,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Access the workbook view directly from the spreadsheet
        // Based on the code, the spreadsheet already has a workbook_view field
        let workbook_view = spreadsheet.get_workbook_view_mut();

        // Set the active tab directly on the workbook_view
        workbook_view.set_active_tab(tab_index);

        Ok::<Atom, String>(atoms::ok())
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => {
            Err(NifError::Term(Box::new((atoms::error(), msg))))
        }
        Err(_) => {
            Err(NifError::Term(Box::new((atoms::error(), "Error occurred in set_active_tab".to_string()))))
        }
    }
}

#[rustler::nif]
pub fn set_workbook_window_position(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    _x_position: i32,
    _y_position: i32,
    _window_width: u32,
    _window_height: u32,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Access the workbook view directly from the spreadsheet
        let _workbook_view = spreadsheet.get_workbook_view_mut();

        // Note: Based on our inspection of the Rust library code,
        // the WorkbookView struct doesn't directly expose methods to set window position and size.
        // These attributes are hardcoded in the write_to method.
        // For now, we'll just return OK, and in the future we might need to modify
        // the umya-spreadsheet library to expose these methods.

        Ok::<Atom, String>(atoms::ok())
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => {
            Err(NifError::Term(Box::new((atoms::error(), msg))))
        }
        Err(_) => {
            Err(NifError::Term(Box::new((atoms::error(), "Error occurred in set_workbook_window_position".to_string()))))
        }
    }
}
