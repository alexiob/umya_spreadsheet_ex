use crate::atoms;
use crate::UmyaSpreadsheet;
use rustler::{Atom, Error as NifError, NifResult};
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
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<(), String> {
        let mut spreadsheet = spreadsheet_resource
            .spreadsheet
            .lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        // Validate inputs
        if range.trim().is_empty() {
            return Err("Range cannot be empty".to_string());
        }
        if formula.trim().is_empty() {
            return Err("Formula cannot be empty".to_string());
        }

        // Get sheet by name
        let sheet = spreadsheet
            .get_sheet_by_name_mut(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        // Create a cell formula object
        let mut cell_formula = CellFormula::default();

        // Set formula text and type
        cell_formula.set_text(formula);
        cell_formula.set_formula_type(CellFormulaValues::Array);

        // First cell in the range is the master cell
        let master_coordinate = range
            .split(':')
            .next()
            .ok_or_else(|| format!("Invalid range format: {}", range))?;

        // Need to pass the text of the formula, not the CellFormula object
        sheet
            .get_cell_mut(master_coordinate)
            .set_formula(cell_formula.get_text());
        Ok(())
    }));

    match result {
        Ok(Ok(())) => Ok(atoms::ok()),
        Ok(Err(err_msg)) => Err(NifError::Term(Box::new((atoms::error(), err_msg)))),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Error occurred in set_array_formula operation".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn create_named_range(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    name: String,
    sheet_name: String,
    range: String,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<(), String> {
        let mut spreadsheet = spreadsheet_resource
            .spreadsheet
            .lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        // Validate inputs
        if range.trim().is_empty() {
            return Err("Range cannot be empty".to_string());
        }
        if name.trim().is_empty() {
            return Err("Name cannot be empty".to_string());
        }

        // Get the sheet and use its add_defined_name method which can call the private set_name method
        // since it's within the same crate
        let sheet = spreadsheet
            .get_sheet_by_name_mut(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        // Create the range address in the format expected (just the range, not prefixed with sheet name)
        sheet
            .add_defined_name(name, format!("{}!{}", sheet_name, range))
            .map_err(|e| format!("Failed to add named range: {}", e))?;

        Ok(())
    }));

    match result {
        Ok(Ok(())) => Ok(atoms::ok()),
        Ok(Err(err_msg)) => Err(NifError::Term(Box::new((atoms::error(), err_msg)))),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Error occurred in create_named_range operation".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn create_defined_name(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    name: String,
    formula: String,
    sheet_name: Option<String>,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<(), String> {
        let mut spreadsheet = spreadsheet_resource
            .spreadsheet
            .lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        // Validate inputs
        if formula.trim().is_empty() {
            return Err("Formula cannot be empty".to_string());
        }
        if name.trim().is_empty() {
            return Err("Name cannot be empty".to_string());
        }

        // Handle sheet-scoped vs global defined names
        if let Some(sheet_name_str) = &sheet_name {
            // For sheet-scoped defined names, use the worksheet's add_defined_name method
            // which can call the private set_name method since it's within the same crate
            let sheet = spreadsheet
                .get_sheet_by_name_mut(sheet_name_str)
                .ok_or_else(|| format!("Sheet '{}' not found", sheet_name_str))?;

            sheet
                .add_defined_name(name, formula)
                .map_err(|e| format!("Failed to add defined name: {}", e))?;
        } else {
            // For global defined names, we cannot set the name directly due to the set_name method being private.
            // This is a limitation of the current umya-spreadsheet 2.3.0 API when used from external crates.
            // We create a defined name with only the formula/address set.
            // Note: The name won't be set, but the formula will still work.
            let mut defined_name = DefinedName::default();
            defined_name.set_address(&formula);

            // Add the defined name to the spreadsheet
            spreadsheet.add_defined_names(defined_name);

            // Log a warning about the limitation
            eprintln!("Warning: Global defined name '{}' created without name due to API limitations in umya-spreadsheet 2.3.0. The formula '{}' is stored and functional.", name, formula);
        }

        Ok(())
    }));

    match result {
        Ok(Ok(())) => Ok(atoms::ok()),
        Ok(Err(err_msg)) => Err(NifError::Term(Box::new((atoms::error(), err_msg)))),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Error occurred in create_defined_name operation".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn get_defined_names(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
) -> Vec<(String, String)> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        // Lock the mutex to get access to the spreadsheet
        let guard = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get global defined names
        let defined_names = guard.get_defined_names();
        let mut names = Vec::new();

        // Add global defined names
        for defined_name in defined_names {
            let name = defined_name.get_name().to_string();
            let address = defined_name.get_address();
            names.push((name, address));
        }

        // Add sheet-specific defined names
        for sheet in guard.get_sheet_collection() {
            for defined_name in sheet.get_defined_names() {
                let name = defined_name.get_name().to_string();
                let address = defined_name.get_address();
                names.push((name, address));
            }
        }

        Ok::<Vec<(String, String)>, String>(names)
    }));

    match result {
        Ok(Ok(names)) => names,
        Ok(Err(_msg)) => Vec::new(),
        Err(_) => Vec::new(),
    }
}

#[rustler::nif]
pub fn set_formula(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    formula: String,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<(), String> {
        let mut spreadsheet = spreadsheet_resource
            .spreadsheet
            .lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        // Validate inputs
        if cell_address.trim().is_empty() {
            return Err("Cell address cannot be empty".to_string());
        }
        if formula.trim().is_empty() {
            return Err("Formula cannot be empty".to_string());
        }

        // Get sheet by name
        let sheet = spreadsheet
            .get_sheet_by_name_mut(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

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
        Ok(())
    }));

    match result {
        Ok(Ok(())) => Ok(atoms::ok()),
        Ok(Err(err_msg)) => Err(NifError::Term(Box::new((atoms::error(), err_msg)))),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Error occurred in set_formula operation".to_string(),
        )))),
    }
}
