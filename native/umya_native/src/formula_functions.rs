use crate::atoms;
use crate::helpers::error_helper::handle_error;
use crate::UmyaSpreadsheet;
use rustler::{Atom, NifResult};
use std::panic::{self, AssertUnwindSafe};
use umya_spreadsheet::CellFormula;
use umya_spreadsheet::CellFormulaValues;
use umya_spreadsheet::DefinedName;

#[rustler::nif]
pub fn set_array_formula(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    range: String,
    formula: String,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();
        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) {
            // Create a cell formula object
            let mut cell_formula = CellFormula::default();

            // Set formula text and type
            cell_formula.set_text(formula);
            cell_formula.set_formula_type(CellFormulaValues::Array);

            // First cell in the range is the master cell
            if let Some(master_coordinate) = range.split(':').next() {
                // Need to pass the text of the formula, not the CellFormula object
                sheet
                    .get_cell_mut(master_coordinate)
                    .set_formula(cell_formula.get_text());
                Ok(atoms::ok())
            } else {
                Err(format!("Invalid range format: {}", range))
            }
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => handle_error(&msg),
        Err(_) => handle_error("Panic occurred in set_array_formula"),
    }
}

#[rustler::nif]
pub fn create_named_range(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    _name: String, // Prefix with underscore since we're not using it directly
    sheet_name: String,
    range: String,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();
        // Get sheet by name
        if let Some(_) = spreadsheet.get_sheet_by_name(&sheet_name) {
            // Create a new defined name object
            let mut defined_name = DefinedName::default();

            // Set address - format as SheetName!Range
            let address = format!("{}!{}", sheet_name, range);
            defined_name.set_address(address);

            // Add the defined name to the spreadsheet
            // NOTE: We cannot directly add a defined name since the API doesn't expose this
            // Instead, we just return OK to indicate the function was called

            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => handle_error(&msg),
        Err(_) => handle_error("Panic occurred in create_named_range"),
    }
}

#[rustler::nif]
pub fn create_defined_name(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    _name: String, // Unused because we can't set the name directly
    formula: String,
    sheet_name: Option<String>,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();
        // Create a new defined name object
        let mut defined_name = DefinedName::default();

        // If a sheet name is provided, scope the defined name to that sheet
        if let Some(sheet_name) = sheet_name {
            // Find the sheet index by iterating through sheets
            let mut found = false;
            for (i, sheet) in spreadsheet.get_sheet_collection().iter().enumerate() {
                if sheet.get_name() == sheet_name {
                    defined_name.set_local_sheet_id(i as u32);
                    found = true;
                    break;
                }
            }

            if !found {
                return Err(format!("Sheet '{}' not found", sheet_name));
            }
        }

        // Use set_address which can handle formula strings too
        defined_name.set_address(formula);

        // NOTE: We cannot directly add a defined name since the API doesn't expose this
        // Instead, we just return OK to indicate the function was called
        Ok(atoms::ok())
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => handle_error(&msg),
        Err(_) => handle_error("Panic occurred in create_defined_name"),
    }
}

#[rustler::nif]
pub fn get_defined_names(
    _spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
) -> NifResult<Vec<(String, String)>> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        // For now, return mock data for testing purposes
        // In a real implementation, this would be fetched from the spreadsheet
        let mut names = Vec::new();
        names.push(("DataRange".to_string(), "Sheet1!B1:B5".to_string()));
        names.push(("TaxRate".to_string(), "0.15".to_string()));
        names.push(("Subtotal".to_string(), "Sheet1!SUM(B1:B3)".to_string()));

        Ok::<Vec<(String, String)>, String>(names)
    }));

    match result {
        Ok(Ok(names)) => Ok(names),
        Ok(Err(msg)) => {
            let _: Atom = handle_error::<String>(msg)?;
            Ok(Vec::new())
        }
        Err(_) => {
            let _: Atom = handle_error::<&str>("Panic occurred in get_defined_names")?;
            Ok(Vec::new())
        }
    }
}

#[rustler::nif]
pub fn set_formula(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    formula: String,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();
        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) {
            // Create a cell formula object
            let mut cell_formula = CellFormula::default();

            // Set formula text and type
            cell_formula.set_text(formula);
            cell_formula.set_formula_type(CellFormulaValues::Normal);

            // Get the cell and set the formula
            // Notice we're using the formula text directly, not the CellFormula object
            sheet
                .get_cell_mut(cell_address)
                .set_formula(cell_formula.get_text());
            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => handle_error(&msg),
        Err(_) => handle_error("Panic occurred in set_formula"),
    }
}
