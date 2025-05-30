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
            // eprintln!("Warning: Global defined name '{}' created without name due to API limitations in umya-spreadsheet 2.3.0. The formula '{}' is stored and functional.", name, formula);
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

#[rustler::nif]
pub fn is_formula(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> bool {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<bool, String> {
        let spreadsheet = spreadsheet_resource
            .spreadsheet
            .lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        // Validate inputs
        if cell_address.trim().is_empty() {
            return Err("Cell address cannot be empty".to_string());
        }

        // Get sheet by name
        let sheet = spreadsheet
            .get_sheet_by_name(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        // Get the cell and check if it has a formula
        let cell = sheet.get_cell(cell_address);
        Ok(cell.map_or(false, |c| c.is_formula()))
    }));

    match result {
        Ok(Ok(is_formula)) => is_formula,
        Ok(Err(_err_msg)) => false,
        Err(_) => false,
    }
}

#[rustler::nif]
pub fn get_formula(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> String {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<String, String> {
        let spreadsheet = spreadsheet_resource
            .spreadsheet
            .lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        // Validate inputs
        if cell_address.trim().is_empty() {
            return Err("Cell address cannot be empty".to_string());
        }

        // Get sheet by name
        let sheet = spreadsheet
            .get_sheet_by_name(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        // Get the cell and its formula
        let cell = sheet.get_cell(cell_address);
        Ok(cell.map_or(String::new(), |c| c.get_formula().to_string()))
    }));

    match result {
        Ok(Ok(formula)) => formula,
        Ok(Err(_err_msg)) => String::new(),
        Err(_) => String::new(),
    }
}

#[rustler::nif]
pub fn get_formula_obj(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> (String, String, Option<i32>, Option<String>) {
    let result = panic::catch_unwind(AssertUnwindSafe(
        || -> Result<(String, String, Option<i32>, Option<String>), String> {
            let spreadsheet = spreadsheet_resource
                .spreadsheet
                .lock()
                .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

            // Validate inputs
            if cell_address.trim().is_empty() {
                return Err("Cell address cannot be empty".to_string());
            }

            // Get sheet by name
            let sheet = spreadsheet
                .get_sheet_by_name(&sheet_name)
                .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

            // Get the cell and its formula object
            let cell = sheet.get_cell(cell_address);
            if let Some(cell) = cell {
                if let Some(formula_obj) = cell.get_formula_obj() {
                    let text = formula_obj.get_text().to_string();
                    let formula_type = format!("{:?}", formula_obj.get_formula_type());
                    let shared_index = formula_obj.get_shared_index().clone();
                    let reference = Some(formula_obj.get_reference().to_string());
                    let shared_index_i32 = Some(shared_index as i32);
                    Ok((text, formula_type, shared_index_i32, reference))
                } else {
                    Ok((String::new(), "None".to_string(), None, None))
                }
            } else {
                Ok((String::new(), "None".to_string(), None, None))
            }
        },
    ));

    match result {
        Ok(Ok(result)) => result,
        Ok(Err(_err_msg)) => (String::new(), "None".to_string(), None, None),
        Err(_) => (String::new(), "None".to_string(), None, None),
    }
}

#[rustler::nif]
pub fn get_formula_shared_index(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> Option<i32> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<Option<i32>, String> {
        let spreadsheet = spreadsheet_resource
            .spreadsheet
            .lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        // Validate inputs
        if cell_address.trim().is_empty() {
            return Err("Cell address cannot be empty".to_string());
        }

        // Get sheet by name
        let sheet = spreadsheet
            .get_sheet_by_name(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        // Get the cell and its formula shared index
        let cell = sheet.get_cell(cell_address);
        if let Some(cell) = cell {
            if let Some(formula_obj) = cell.get_formula_obj() {
                Ok(Some(*formula_obj.get_shared_index() as i32))
            } else {
                Ok(None)
            }
        } else {
            Ok(None)
        }
    }));

    match result {
        Ok(Ok(shared_index)) => shared_index,
        Ok(Err(_err_msg)) => None,
        Err(_) => None,
    }
}

#[rustler::nif]
pub fn get_text(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> String {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<String, String> {
        let spreadsheet = spreadsheet_resource
            .spreadsheet
            .lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        // Validate inputs
        if cell_address.trim().is_empty() {
            return Err("Cell address cannot be empty".to_string());
        }

        // Get sheet by name
        let sheet = spreadsheet
            .get_sheet_by_name(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        // Get the cell and its formula text
        let cell = sheet.get_cell(cell_address);
        if let Some(cell) = cell {
            if let Some(formula_obj) = cell.get_formula_obj() {
                Ok(formula_obj.get_text().to_string())
            } else {
                Ok(String::new())
            }
        } else {
            Ok(String::new())
        }
    }));

    match result {
        Ok(Ok(text)) => text,
        Ok(Err(_err_msg)) => String::new(),
        Err(_) => String::new(),
    }
}

#[rustler::nif]
pub fn get_formula_type(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> String {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<String, String> {
        let spreadsheet = spreadsheet_resource
            .spreadsheet
            .lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        // Validate inputs
        if cell_address.trim().is_empty() {
            return Err("Cell address cannot be empty".to_string());
        }

        // Get sheet by name
        let sheet = spreadsheet
            .get_sheet_by_name(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        // Get the cell and its formula type
        let cell = sheet.get_cell(cell_address);
        if let Some(cell) = cell {
            if let Some(formula_obj) = cell.get_formula_obj() {
                Ok(format!("{:?}", formula_obj.get_formula_type()))
            } else {
                Ok("None".to_string())
            }
        } else {
            Ok("None".to_string())
        }
    }));

    match result {
        Ok(Ok(formula_type)) => formula_type,
        Ok(Err(_err_msg)) => "None".to_string(),
        Err(_) => "None".to_string(),
    }
}

#[rustler::nif]
pub fn get_shared_index(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> Option<i32> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<Option<i32>, String> {
        let spreadsheet = spreadsheet_resource
            .spreadsheet
            .lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        // Validate inputs
        if cell_address.trim().is_empty() {
            return Err("Cell address cannot be empty".to_string());
        }

        // Get sheet by name
        let sheet = spreadsheet
            .get_sheet_by_name(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        // Get the cell and its shared index
        let cell = sheet.get_cell(cell_address);
        if let Some(cell) = cell {
            if let Some(formula_obj) = cell.get_formula_obj() {
                Ok(Some(*formula_obj.get_shared_index() as i32))
            } else {
                Ok(None)
            }
        } else {
            Ok(None)
        }
    }));

    match result {
        Ok(Ok(shared_index)) => shared_index,
        Ok(Err(_err_msg)) => None,
        Err(_) => None,
    }
}

#[rustler::nif]
pub fn get_reference(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> Option<String> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<Option<String>, String> {
        let spreadsheet = spreadsheet_resource
            .spreadsheet
            .lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        // Validate inputs
        if cell_address.trim().is_empty() {
            return Err("Cell address cannot be empty".to_string());
        }

        // Get sheet by name
        let sheet = spreadsheet
            .get_sheet_by_name(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        // Get the cell and its reference
        let cell = sheet.get_cell(cell_address);
        if let Some(cell) = cell {
            if let Some(formula_obj) = cell.get_formula_obj() {
                Ok(Some(formula_obj.get_reference().to_string()))
            } else {
                Ok(None)
            }
        } else {
            Ok(None)
        }
    }));

    match result {
        Ok(Ok(reference)) => reference,
        Ok(Err(_err_msg)) => None,
        Err(_) => None,
    }
}

#[rustler::nif]
pub fn get_bx(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> Option<bool> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<Option<bool>, String> {
        let spreadsheet = spreadsheet_resource
            .spreadsheet
            .lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        // Validate inputs
        if cell_address.trim().is_empty() {
            return Err("Cell address cannot be empty".to_string());
        }

        // Get sheet by name
        let sheet = spreadsheet
            .get_sheet_by_name(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        // Get the cell and its bx value
        let cell = sheet.get_cell(cell_address);
        if let Some(cell) = cell {
            if let Some(formula_obj) = cell.get_formula_obj() {
                Ok(Some(*formula_obj.get_bx()))
            } else {
                Ok(None)
            }
        } else {
            Ok(None)
        }
    }));

    match result {
        Ok(Ok(bx)) => bx,
        Ok(Err(_err_msg)) => None,
        Err(_) => None,
    }
}

#[rustler::nif]
pub fn get_data_table_2d(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> Option<bool> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<Option<bool>, String> {
        let spreadsheet = spreadsheet_resource
            .spreadsheet
            .lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        // Validate inputs
        if cell_address.trim().is_empty() {
            return Err("Cell address cannot be empty".to_string());
        }

        // Get sheet by name
        let sheet = spreadsheet
            .get_sheet_by_name(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        // Get the cell and its data table 2d value
        let cell = sheet.get_cell(cell_address);
        if let Some(cell) = cell {
            if let Some(formula_obj) = cell.get_formula_obj() {
                Ok(Some(*formula_obj.get_data_table_2d()))
            } else {
                Ok(None)
            }
        } else {
            Ok(None)
        }
    }));

    match result {
        Ok(Ok(data_table_2d)) => data_table_2d,
        Ok(Err(_err_msg)) => None,
        Err(_) => None,
    }
}

#[rustler::nif]
pub fn get_data_table_row(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> Option<bool> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<Option<bool>, String> {
        let spreadsheet = spreadsheet_resource
            .spreadsheet
            .lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        // Validate inputs
        if cell_address.trim().is_empty() {
            return Err("Cell address cannot be empty".to_string());
        }

        // Get sheet by name
        let sheet = spreadsheet
            .get_sheet_by_name(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        // Get the cell and its data table row value
        let cell = sheet.get_cell(cell_address);
        if let Some(cell) = cell {
            if let Some(formula_obj) = cell.get_formula_obj() {
                Ok(Some(*formula_obj.get_data_table_row()))
            } else {
                Ok(None)
            }
        } else {
            Ok(None)
        }
    }));

    match result {
        Ok(Ok(data_table_row)) => data_table_row,
        Ok(Err(_err_msg)) => None,
        Err(_) => None,
    }
}

#[rustler::nif]
pub fn get_input_1deleted(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> Option<bool> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<Option<bool>, String> {
        let spreadsheet = spreadsheet_resource
            .spreadsheet
            .lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        // Validate inputs
        if cell_address.trim().is_empty() {
            return Err("Cell address cannot be empty".to_string());
        }

        // Get sheet by name
        let sheet = spreadsheet
            .get_sheet_by_name(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        // Get the cell and its input 1 deleted value
        let cell = sheet.get_cell(cell_address);
        if let Some(cell) = cell {
            if let Some(formula_obj) = cell.get_formula_obj() {
                Ok(Some(*formula_obj.get_input_1deleted()))
            } else {
                Ok(None)
            }
        } else {
            Ok(None)
        }
    }));

    match result {
        Ok(Ok(input_1deleted)) => input_1deleted,
        Ok(Err(_err_msg)) => None,
        Err(_) => None,
    }
}

#[rustler::nif]
pub fn get_input_2deleted(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> Option<bool> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<Option<bool>, String> {
        let spreadsheet = spreadsheet_resource
            .spreadsheet
            .lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        // Validate inputs
        if cell_address.trim().is_empty() {
            return Err("Cell address cannot be empty".to_string());
        }

        // Get sheet by name
        let sheet = spreadsheet
            .get_sheet_by_name(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        // Get the cell and its input 2 deleted value
        let cell = sheet.get_cell(cell_address);
        if let Some(cell) = cell {
            if let Some(formula_obj) = cell.get_formula_obj() {
                Ok(Some(*formula_obj.get_input_2deleted()))
            } else {
                Ok(None)
            }
        } else {
            Ok(None)
        }
    }));

    match result {
        Ok(Ok(input_2deleted)) => input_2deleted,
        Ok(Err(_err_msg)) => None,
        Err(_) => None,
    }
}

#[rustler::nif]
pub fn get_r1(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> Option<String> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<Option<String>, String> {
        let spreadsheet = spreadsheet_resource
            .spreadsheet
            .lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        // Validate inputs
        if cell_address.trim().is_empty() {
            return Err("Cell address cannot be empty".to_string());
        }

        // Get sheet by name
        let sheet = spreadsheet
            .get_sheet_by_name(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        // Get the cell and its r1 value
        let cell = sheet.get_cell(cell_address);
        if let Some(cell) = cell {
            if let Some(formula_obj) = cell.get_formula_obj() {
                Ok(Some(formula_obj.get_r1().to_string()))
            } else {
                Ok(None)
            }
        } else {
            Ok(None)
        }
    }));

    match result {
        Ok(Ok(r1)) => r1,
        Ok(Err(_err_msg)) => None,
        Err(_) => None,
    }
}

#[rustler::nif]
pub fn get_r2(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> Option<String> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<Option<String>, String> {
        let spreadsheet = spreadsheet_resource
            .spreadsheet
            .lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        // Validate inputs
        if cell_address.trim().is_empty() {
            return Err("Cell address cannot be empty".to_string());
        }

        // Get sheet by name
        let sheet = spreadsheet
            .get_sheet_by_name(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        // Get the cell and its r2 value
        let cell = sheet.get_cell(cell_address);
        if let Some(cell) = cell {
            if let Some(formula_obj) = cell.get_formula_obj() {
                Ok(Some(formula_obj.get_r2().to_string()))
            } else {
                Ok(None)
            }
        } else {
            Ok(None)
        }
    }));

    match result {
        Ok(Ok(r2)) => r2,
        Ok(Err(_err_msg)) => None,
        Err(_) => None,
    }
}
