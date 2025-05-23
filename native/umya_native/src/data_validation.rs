use rustler::{Atom, ResourceArc};
use umya_spreadsheet::{DataValidationOperatorValues, DataValidationValues};

use crate::atoms;
use crate::UmyaSpreadsheet;

/// Add a simple list data validation to a cell or range
#[rustler::nif]
pub fn add_list_validation(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_range: String,
    list_items: Vec<String>,
    allow_blank: bool,
    error_title: Option<String>,
    error_message: Option<String>,
    prompt_title: Option<String>,
    prompt_message: Option<String>,
) -> Result<Atom, Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Create the data validation
            let mut validation = umya_spreadsheet::DataValidation::default();

            // Set basic properties
            validation.set_type(DataValidationValues::List);
            validation.set_allow_blank(allow_blank);

            // Create and set sequence of references
            let mut seq_refs = umya_spreadsheet::SequenceOfReferences::default();
            seq_refs.set_sqref(&cell_range);
            validation.set_sequence_of_references(seq_refs);

            // Set the formula1 (list source)
            let formula = list_items.join(",");
            validation.set_formula1(format!("\"{}\"", formula));

            // Set error title and message if provided
            if let Some(title) = error_title {
                validation.set_error_title(title);
            }

            if let Some(msg) = error_message {
                validation.set_error_message(msg);
                validation.set_show_error_message(true);
            }

            // Set prompt title and message if provided
            if let Some(title) = prompt_title {
                validation.set_prompt_title(title);
            }

            if let Some(msg) = prompt_message {
                validation.set_prompt(msg);
                validation.set_show_input_message(true);
            }

            // Add the validation to the sheet
            if let Some(validations) = sheet.get_data_validations_mut() {
                validations.add_data_validation_list(validation);
            } else {
                let mut data_validations = umya_spreadsheet::DataValidations::default();
                data_validations.add_data_validation_list(validation);
                sheet.set_data_validations(data_validations);
            }

            Ok(atoms::ok())
        }
        None => Err(atoms::not_found()),
    }
}

/// Add a number data validation to a cell or range
#[rustler::nif]
pub fn add_number_validation(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_range: String,
    operator: String,
    value1: String,
    value2: Option<String>,
    allow_blank: bool,
    error_title: Option<String>,
    error_message: Option<String>,
    prompt_title: Option<String>,
    prompt_message: Option<String>,
) -> Result<Atom, Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    // Convert operator string to Rust enum
    let op = match operator.as_str() {
        "between" => DataValidationOperatorValues::Between,
        "notBetween" | "not_between" => DataValidationOperatorValues::NotBetween,
        "equal" => DataValidationOperatorValues::Equal,
        "notEqual" | "not_equal" => DataValidationOperatorValues::NotEqual,
        "greaterThan" | "greater_than" => DataValidationOperatorValues::GreaterThan,
        "lessThan" | "less_than" => DataValidationOperatorValues::LessThan,
        "greaterThanOrEqual" | "greater_than_or_equal" => {
            DataValidationOperatorValues::GreaterThanOrEqual
        }
        "lessThanOrEqual" | "less_than_or_equal" => DataValidationOperatorValues::LessThanOrEqual,
        _ => return Err(atoms::error()),
    };

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Create the data validation
            let mut validation = umya_spreadsheet::DataValidation::default();

            // Set basic properties
            validation.set_type(DataValidationValues::Decimal);
            validation.set_operator(op);
            validation.set_allow_blank(allow_blank);

            // Create and set sequence of references
            let mut seq_refs = umya_spreadsheet::SequenceOfReferences::default();
            seq_refs.set_sqref(&cell_range);
            validation.set_sequence_of_references(seq_refs);

            // Set the formula values
            validation.set_formula1(value1);
            if let Some(val2) = value2 {
                validation.set_formula2(val2);
            }

            // Set error title and message if provided
            if let Some(title) = error_title {
                validation.set_error_title(title);
            }

            if let Some(msg) = error_message {
                validation.set_error_message(msg);
                validation.set_show_error_message(true);
            }

            // Set prompt title and message if provided
            if let Some(title) = prompt_title {
                validation.set_prompt_title(title);
            }

            if let Some(msg) = prompt_message {
                validation.set_prompt(msg);
                validation.set_show_input_message(true);
            }

            // Add the validation to the sheet
            if let Some(validations) = sheet.get_data_validations_mut() {
                validations.add_data_validation_list(validation);
            } else {
                let mut data_validations = umya_spreadsheet::DataValidations::default();
                data_validations.add_data_validation_list(validation);
                sheet.set_data_validations(data_validations);
            }

            Ok(atoms::ok())
        }
        None => Err(atoms::not_found()),
    }
}

/// Add a date data validation to a cell or range
#[rustler::nif]
pub fn add_date_validation(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_range: String,
    operator: String,
    date1: String,
    date2: Option<String>,
    allow_blank: bool,
    error_title: Option<String>,
    error_message: Option<String>,
    prompt_title: Option<String>,
    prompt_message: Option<String>,
) -> Result<Atom, Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    // Convert operator string to Rust enum
    let op = match operator.as_str() {
        "between" => DataValidationOperatorValues::Between,
        "notBetween" | "not_between" => DataValidationOperatorValues::NotBetween,
        "equal" => DataValidationOperatorValues::Equal,
        "notEqual" | "not_equal" => DataValidationOperatorValues::NotEqual,
        "greaterThan" | "greater_than" => DataValidationOperatorValues::GreaterThan,
        "lessThan" | "less_than" => DataValidationOperatorValues::LessThan,
        "greaterThanOrEqual" | "greater_than_or_equal" => {
            DataValidationOperatorValues::GreaterThanOrEqual
        }
        "lessThanOrEqual" | "less_than_or_equal" => DataValidationOperatorValues::LessThanOrEqual,
        _ => return Err(atoms::error()),
    };

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Create the data validation
            let mut validation = umya_spreadsheet::DataValidation::default();

            // Set basic properties
            validation.set_type(DataValidationValues::Date);
            validation.set_operator(op);
            validation.set_allow_blank(allow_blank);

            // Create and set sequence of references
            let mut seq_refs = umya_spreadsheet::SequenceOfReferences::default();
            seq_refs.set_sqref(&cell_range);
            validation.set_sequence_of_references(seq_refs);

            // Set the formula values
            validation.set_formula1(date1);
            if let Some(date2) = date2 {
                validation.set_formula2(date2);
            }

            // Set error title and message if provided
            if let Some(title) = error_title {
                validation.set_error_title(title);
            }

            if let Some(msg) = error_message {
                validation.set_error_message(msg);
                validation.set_show_error_message(true);
            }

            // Set prompt title and message if provided
            if let Some(title) = prompt_title {
                validation.set_prompt_title(title);
            }

            if let Some(msg) = prompt_message {
                validation.set_prompt(msg);
                validation.set_show_input_message(true);
            }

            // Add the validation to the sheet
            if let Some(validations) = sheet.get_data_validations_mut() {
                validations.add_data_validation_list(validation);
            } else {
                let mut data_validations = umya_spreadsheet::DataValidations::default();
                data_validations.add_data_validation_list(validation);
                sheet.set_data_validations(data_validations);
            }

            Ok(atoms::ok())
        }
        None => Err(atoms::not_found()),
    }
}

/// Add a text length data validation to a cell or range
#[rustler::nif]
pub fn add_text_length_validation(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_range: String,
    operator: String,
    length1: i32,
    length2: Option<i32>,
    allow_blank: bool,
    error_title: Option<String>,
    error_message: Option<String>,
    prompt_title: Option<String>,
    prompt_message: Option<String>,
) -> Result<Atom, Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    // Convert operator string to Rust enum
    let op = match operator.as_str() {
        "between" => DataValidationOperatorValues::Between,
        "notBetween" | "not_between" => DataValidationOperatorValues::NotBetween,
        "equal" => DataValidationOperatorValues::Equal,
        "notEqual" | "not_equal" => DataValidationOperatorValues::NotEqual,
        "greaterThan" | "greater_than" => DataValidationOperatorValues::GreaterThan,
        "lessThan" | "less_than" => DataValidationOperatorValues::LessThan,
        "greaterThanOrEqual" | "greater_than_or_equal" => {
            DataValidationOperatorValues::GreaterThanOrEqual
        }
        "lessThanOrEqual" | "less_than_or_equal" => DataValidationOperatorValues::LessThanOrEqual,
        _ => return Err(atoms::error()),
    };

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Create the data validation
            let mut validation = umya_spreadsheet::DataValidation::default();

            // Set basic properties
            validation.set_type(DataValidationValues::TextLength);
            validation.set_operator(op);
            validation.set_allow_blank(allow_blank);

            // Create and set sequence of references
            let mut seq_refs = umya_spreadsheet::SequenceOfReferences::default();
            seq_refs.set_sqref(&cell_range);
            validation.set_sequence_of_references(seq_refs);

            // Set the formula values
            validation.set_formula1(length1.to_string());
            if let Some(len2) = length2 {
                validation.set_formula2(len2.to_string());
            }

            // Set error title and message if provided
            if let Some(title) = error_title {
                validation.set_error_title(title);
            }

            if let Some(msg) = error_message {
                validation.set_error_message(msg);
                validation.set_show_error_message(true);
            }

            // Set prompt title and message if provided
            if let Some(title) = prompt_title {
                validation.set_prompt_title(title);
            }

            if let Some(msg) = prompt_message {
                validation.set_prompt(msg);
                validation.set_show_input_message(true);
            }

            // Add the validation to the sheet
            if let Some(validations) = sheet.get_data_validations_mut() {
                validations.add_data_validation_list(validation);
            } else {
                let mut data_validations = umya_spreadsheet::DataValidations::default();
                data_validations.add_data_validation_list(validation);
                sheet.set_data_validations(data_validations);
            }

            Ok(atoms::ok())
        }
        None => Err(atoms::not_found()),
    }
}

/// Add a custom formula validation to a cell or range
#[rustler::nif]
pub fn add_custom_validation(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_range: String,
    formula: String,
    allow_blank: bool,
    error_title: Option<String>,
    error_message: Option<String>,
    prompt_title: Option<String>,
    prompt_message: Option<String>,
) -> Result<Atom, Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Create the data validation
            let mut validation = umya_spreadsheet::DataValidation::default();

            // Set basic properties
            validation.set_type(DataValidationValues::Custom);
            validation.set_allow_blank(allow_blank);

            // Create and set sequence of references
            let mut seq_refs = umya_spreadsheet::SequenceOfReferences::default();
            seq_refs.set_sqref(&cell_range);
            validation.set_sequence_of_references(seq_refs);

            // Set the formula
            validation.set_formula1(formula);

            // Set error title and message if provided
            if let Some(title) = error_title {
                validation.set_error_title(title);
            }

            if let Some(msg) = error_message {
                validation.set_error_message(msg);
                validation.set_show_error_message(true);
            }

            // Set prompt title and message if provided
            if let Some(title) = prompt_title {
                validation.set_prompt_title(title);
            }

            if let Some(msg) = prompt_message {
                validation.set_prompt(msg);
                validation.set_show_input_message(true);
            }

            // Add the validation to the sheet
            if let Some(validations) = sheet.get_data_validations_mut() {
                validations.add_data_validation_list(validation);
            } else {
                let mut data_validations = umya_spreadsheet::DataValidations::default();
                data_validations.add_data_validation_list(validation);
                sheet.set_data_validations(data_validations);
            }

            Ok(atoms::ok())
        }
        None => Err(atoms::not_found()),
    }
}

/// Remove data validations from a cell range
#[rustler::nif]
pub fn remove_data_validation(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_range: String,
) -> Result<Atom, Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Get current validations
            if let Some(validations) = sheet.get_data_validations_mut() {
                // Filter out validations that match the cell range
                let mut range_to_remove = umya_spreadsheet::SequenceOfReferences::default();
                range_to_remove.set_sqref(&cell_range);

                // Since we can't directly remove validations from the list, we'll collect
                // the validations to keep and then replace the original list
                let mut validations_to_keep = Vec::new();

                for validation in validations.get_data_validation_list().iter() {
                    if validation.get_sequence_of_references().get_sqref()
                        != range_to_remove.get_sqref()
                    {
                        validations_to_keep.push(validation.clone());
                    }
                }

                // Replace the original list with our filtered list
                validations.set_data_validation_list(validations_to_keep);
            }

            Ok(atoms::ok())
        }
        None => Err(atoms::not_found()),
    }
}
