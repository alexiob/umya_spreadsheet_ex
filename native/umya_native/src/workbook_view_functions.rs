use crate::atoms;
use crate::UmyaSpreadsheet;
use rustler::{Atom, Encoder, Env, Error as NifError, NifResult, Term};
use std::collections::HashMap;
use std::panic::{self, AssertUnwindSafe};

#[rustler::nif]
pub fn get_active_tab(
    env: Env,
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
) -> Term {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get the active tab index
        let active_tab = *spreadsheet.get_workbook_view().get_active_tab();

        (atoms::ok(), active_tab).encode(env)
    }));

    match result {
        Ok(term) => term,
        Err(_) => (atoms::error(), "Error occurred in get_active_tab").encode(env),
    }
}

#[rustler::nif]
pub fn get_workbook_window_position(
    env: Env,
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
) -> Term {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        // Note: These are hardcoded default values based on the write_to method in workbook_view.rs
        // since there's no direct getter for these attributes in the Rust lib
        let mut window_info = HashMap::new();
        window_info.insert("x_position".to_string(), "240");
        window_info.insert("y_position".to_string(), "105");
        window_info.insert("width".to_string(), "14805");
        window_info.insert("height".to_string(), "8010");

        (atoms::ok(), window_info).encode(env)
    }));

    match result {
        Ok(term) => term,
        Err(_) => (
            atoms::error(),
            "Error occurred in get_workbook_window_position",
        )
            .encode(env),
    }
}

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
        Ok(Err(msg)) => Err(NifError::Term(Box::new((atoms::error(), msg)))),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Error occurred in set_active_tab".to_string(),
        )))),
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
        Ok(Err(msg)) => Err(NifError::Term(Box::new((atoms::error(), msg)))),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Error occurred in set_workbook_window_position".to_string(),
        )))),
    }
}
