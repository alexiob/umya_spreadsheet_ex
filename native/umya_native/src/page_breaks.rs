use crate::atoms;
use crate::UmyaSpreadsheet;
use rustler::{Atom, Error as NifError, NifResult, ResourceArc};
use umya_spreadsheet::Break;

/// Add a row page break at the specified row number.
#[rustler::nif]
pub fn add_row_page_break(
    spreadsheet: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    row_number: u32,
    manual: bool,
) -> NifResult<Atom> {
    let result = std::panic::catch_unwind(|| {
        let mut workbook = spreadsheet.spreadsheet.lock().unwrap();
        if let Some(worksheet) = workbook.get_sheet_by_name_mut(&sheet_name) {
            let mut page_break = Break::default();
            page_break.set_id(row_number);
            page_break.set_manual_page_break(manual);

            worksheet.get_row_breaks_mut().add_break_list(page_break);
            Ok(atoms::ok())
        } else {
            Err(NifError::Term(Box::new((
                atoms::error(),
                "sheet_not_found".to_string(),
            ))))
        }
    });

    match result {
        Ok(r) => r,
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "unknown_error".to_string(),
        )))),
    }
}

/// Add a column page break at the specified column number.
#[rustler::nif]
pub fn add_column_page_break(
    spreadsheet: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    column_number: u32,
    manual: bool,
) -> NifResult<Atom> {
    let result = std::panic::catch_unwind(|| {
        let mut workbook = spreadsheet.spreadsheet.lock().unwrap();
        if let Some(worksheet) = workbook.get_sheet_by_name_mut(&sheet_name) {
            let mut page_break = Break::default();
            page_break.set_id(column_number);
            page_break.set_manual_page_break(manual);

            worksheet.get_column_breaks_mut().add_break_list(page_break);
            Ok(atoms::ok())
        } else {
            Err(NifError::Term(Box::new((
                atoms::error(),
                "sheet_not_found".to_string(),
            ))))
        }
    });

    match result {
        Ok(r) => r,
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "unknown_error".to_string(),
        )))),
    }
}

/// Remove a row page break at the specified row number.
#[rustler::nif]
pub fn remove_row_page_break(
    spreadsheet: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    row_number: u32,
) -> NifResult<Atom> {
    let result = std::panic::catch_unwind(|| {
        let mut workbook = spreadsheet.spreadsheet.lock().unwrap();
        if let Some(worksheet) = workbook.get_sheet_by_name_mut(&sheet_name) {
            let row_breaks = worksheet.get_row_breaks_mut();
            let break_list = row_breaks.get_break_list_mut();
            break_list.retain(|page_break| *page_break.get_id() != row_number);
            Ok(atoms::ok())
        } else {
            Err(NifError::Term(Box::new((
                atoms::error(),
                "sheet_not_found".to_string(),
            ))))
        }
    });

    match result {
        Ok(r) => r,
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "unknown_error".to_string(),
        )))),
    }
}

/// Remove a column page break at the specified column number.
#[rustler::nif]
pub fn remove_column_page_break(
    spreadsheet: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    column_number: u32,
) -> NifResult<Atom> {
    let result = std::panic::catch_unwind(|| {
        let mut workbook = spreadsheet.spreadsheet.lock().unwrap();
        if let Some(worksheet) = workbook.get_sheet_by_name_mut(&sheet_name) {
            let column_breaks = worksheet.get_column_breaks_mut();
            let break_list = column_breaks.get_break_list_mut();
            break_list.retain(|page_break| *page_break.get_id() != column_number);
            Ok(atoms::ok())
        } else {
            Err(NifError::Term(Box::new((
                atoms::error(),
                "sheet_not_found".to_string(),
            ))))
        }
    });

    match result {
        Ok(r) => r,
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "unknown_error".to_string(),
        )))),
    }
}

/// Get all row page breaks for the specified sheet.
#[rustler::nif]
pub fn get_row_page_breaks(
    spreadsheet: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> NifResult<Vec<(u32, bool)>> {
    let result = std::panic::catch_unwind(|| {
        let workbook = spreadsheet.spreadsheet.lock().unwrap();
        if let Some(worksheet) = workbook.get_sheet_by_name(&sheet_name) {
            let row_breaks = worksheet.get_row_breaks();
            let breaks: Vec<(u32, bool)> = row_breaks
                .get_break_list()
                .iter()
                .map(|page_break| (*page_break.get_id(), *page_break.get_manual_page_break()))
                .collect();
            Ok(breaks)
        } else {
            Err(NifError::Term(Box::new((
                atoms::error(),
                "sheet_not_found".to_string(),
            ))))
        }
    });

    match result {
        Ok(r) => r,
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "unknown_error".to_string(),
        )))),
    }
}

/// Get all column page breaks for the specified sheet.
#[rustler::nif]
pub fn get_column_page_breaks(
    spreadsheet: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> NifResult<Vec<(u32, bool)>> {
    let result = std::panic::catch_unwind(|| {
        let workbook = spreadsheet.spreadsheet.lock().unwrap();
        if let Some(worksheet) = workbook.get_sheet_by_name(&sheet_name) {
            let column_breaks = worksheet.get_column_breaks();
            let breaks: Vec<(u32, bool)> = column_breaks
                .get_break_list()
                .iter()
                .map(|page_break| (*page_break.get_id(), *page_break.get_manual_page_break()))
                .collect();
            Ok(breaks)
        } else {
            Err(NifError::Term(Box::new((
                atoms::error(),
                "sheet_not_found".to_string(),
            ))))
        }
    });

    match result {
        Ok(r) => r,
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "unknown_error".to_string(),
        )))),
    }
}

/// Clear all row page breaks for the specified sheet.
#[rustler::nif]
pub fn clear_row_page_breaks(
    spreadsheet: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> NifResult<Atom> {
    let result = std::panic::catch_unwind(|| {
        let mut workbook = spreadsheet.spreadsheet.lock().unwrap();
        if let Some(worksheet) = workbook.get_sheet_by_name_mut(&sheet_name) {
            let row_breaks = worksheet.get_row_breaks_mut();
            row_breaks.get_break_list_mut().clear();
            Ok(atoms::ok())
        } else {
            Err(NifError::Term(Box::new((
                atoms::error(),
                "sheet_not_found".to_string(),
            ))))
        }
    });

    match result {
        Ok(r) => r,
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "unknown_error".to_string(),
        )))),
    }
}

/// Clear all column page breaks for the specified sheet.
#[rustler::nif]
pub fn clear_column_page_breaks(
    spreadsheet: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> NifResult<Atom> {
    let result = std::panic::catch_unwind(|| {
        let mut workbook = spreadsheet.spreadsheet.lock().unwrap();
        if let Some(worksheet) = workbook.get_sheet_by_name_mut(&sheet_name) {
            let column_breaks = worksheet.get_column_breaks_mut();
            column_breaks.get_break_list_mut().clear();
            Ok(atoms::ok())
        } else {
            Err(NifError::Term(Box::new((
                atoms::error(),
                "sheet_not_found".to_string(),
            ))))
        }
    });

    match result {
        Ok(r) => r,
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "unknown_error".to_string(),
        )))),
    }
}

/// Check if a row page break exists at the specified row number.
#[rustler::nif]
pub fn has_row_page_break(
    spreadsheet: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    row_number: u32,
) -> NifResult<bool> {
    let result = std::panic::catch_unwind(|| {
        let workbook = spreadsheet.spreadsheet.lock().unwrap();
        if let Some(worksheet) = workbook.get_sheet_by_name(&sheet_name) {
            let row_breaks = worksheet.get_row_breaks();
            let has_break = row_breaks
                .get_break_list()
                .iter()
                .any(|page_break| *page_break.get_id() == row_number);
            Ok(has_break)
        } else {
            Err(NifError::Term(Box::new((
                atoms::error(),
                "sheet_not_found".to_string(),
            ))))
        }
    });

    match result {
        Ok(r) => r,
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "unknown_error".to_string(),
        )))),
    }
}

/// Check if a column page break exists at the specified column number.
#[rustler::nif]
pub fn has_column_page_break(
    spreadsheet: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    column_number: u32,
) -> NifResult<bool> {
    let result = std::panic::catch_unwind(|| {
        let workbook = spreadsheet.spreadsheet.lock().unwrap();
        if let Some(worksheet) = workbook.get_sheet_by_name(&sheet_name) {
            let column_breaks = worksheet.get_column_breaks();
            let has_break = column_breaks
                .get_break_list()
                .iter()
                .any(|page_break| *page_break.get_id() == column_number);
            Ok(has_break)
        } else {
            Err(NifError::Term(Box::new((
                atoms::error(),
                "sheet_not_found".to_string(),
            ))))
        }
    });

    match result {
        Ok(r) => r,
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "unknown_error".to_string(),
        )))),
    }
}
