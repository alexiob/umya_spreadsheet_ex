use std::panic::{self, AssertUnwindSafe};
use rustler::{Atom, ResourceArc};
use umya_spreadsheet::DataValidationValues;

use crate::atoms;
use crate::UmyaSpreadsheet;

// Helper function to map DataValidationValues to string
fn validation_type_to_string(val_type: &DataValidationValues) -> String {
    match val_type {
        DataValidationValues::List => "list".to_string(),
        DataValidationValues::Decimal => "decimal".to_string(),
        DataValidationValues::Date => "date".to_string(),
        DataValidationValues::Time => "time".to_string(),
        DataValidationValues::TextLength => "textLength".to_string(),
        DataValidationValues::Custom => "custom".to_string(),
        DataValidationValues::Whole => "whole".to_string(),
        _ => "unknown".to_string(),
    }
}

// Helper function to convert operator to string
fn operator_to_string(op: &umya_spreadsheet::DataValidationOperatorValues) -> String {
    match op {
        umya_spreadsheet::DataValidationOperatorValues::Between => "between".to_string(),
        umya_spreadsheet::DataValidationOperatorValues::NotBetween => "not_between".to_string(),
        umya_spreadsheet::DataValidationOperatorValues::Equal => "equal".to_string(),
        umya_spreadsheet::DataValidationOperatorValues::NotEqual => "not_equal".to_string(),
        umya_spreadsheet::DataValidationOperatorValues::GreaterThan => "greater_than".to_string(),
        umya_spreadsheet::DataValidationOperatorValues::LessThan => "less_than".to_string(),
        umya_spreadsheet::DataValidationOperatorValues::GreaterThanOrEqual => "greater_than_or_equal".to_string(),
        umya_spreadsheet::DataValidationOperatorValues::LessThanOrEqual => "less_than_or_equal".to_string(),
    }
}

/// Get all data validation rules for a specific sheet or range
#[rustler::nif(name = "get_data_validations")]
pub fn get_data_validations_nif<'a>(
    env: rustler::Env<'a>,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_range: Option<String>,
) -> Result<Vec<rustler::Term<'a>>, (Atom, String)> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<Vec<rustler::Term<'a>>, String> {
        let spreadsheet_guard = resource.spreadsheet.lock().unwrap();
        let sheet = spreadsheet_guard
            .get_sheet_by_name(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        let mut validations = Vec::new();

        if let Some(data_validations) = sheet.get_data_validations() {
            for validation in data_validations.get_data_validation_list() {
                // If a specific cell range filter is provided
                if let Some(ref range_filter) = cell_range {
                    if validation.get_sequence_of_references().get_sqref() != *range_filter {
                        continue;
                    }
                }

                let mut map = rustler::types::map::map_new(env);

                // Common properties for all validation types
                map = map.map_put(atoms::range(), validation.get_sequence_of_references().get_sqref().to_string()).ok().unwrap();
                
                // Type-specific properties
                let validation_type = validation_type_to_string(validation.get_type());
                map = map.map_put(atoms::rule_type(), rustler::types::atom::Atom::from_str(env, &validation_type).unwrap()).ok().unwrap();
                
                // Only include operator if applicable for this validation type
                if *validation.get_type() != DataValidationValues::List && *validation.get_type() != DataValidationValues::Custom {
                    let op_str = operator_to_string(validation.get_operator());
                    map = map.map_put(atoms::operator(), rustler::types::atom::Atom::from_str(env, &op_str).unwrap()).ok().unwrap();
                }

                // Formula 1 is present in all validation types
                if !validation.get_formula1().is_empty() {
                    map = map.map_put(atoms::formula1(), validation.get_formula1().to_string()).ok().unwrap();
                }

                // Formula 2 is only present in "between" and "notBetween" operators
                if !validation.get_formula2().is_empty() {
                    map = map.map_put(atoms::formula2(), validation.get_formula2().to_string()).ok().unwrap();
                }

                // Include error message details if present
                if *validation.get_show_error_message() {
                    map = map.map_put(atoms::show_error_message(), validation.get_show_error_message()).ok().unwrap();
                    if !validation.get_error_title().is_empty() {
                        map = map.map_put(atoms::error_title(), validation.get_error_title().to_string()).ok().unwrap();
                    }
                    if !validation.get_error_message().is_empty() {
                        map = map.map_put(atoms::error_message(), validation.get_error_message().to_string()).ok().unwrap();
                    }
                }

                // Include input message details if present
                if *validation.get_show_input_message() {
                    map = map.map_put(atoms::show_input_message(), validation.get_show_input_message()).ok().unwrap();
                    if !validation.get_prompt_title().is_empty() {
                        map = map.map_put(atoms::prompt_title(), validation.get_prompt_title().to_string()).ok().unwrap();
                    }
                    if !validation.get_prompt().is_empty() {
                        map = map.map_put(atoms::prompt_message(), validation.get_prompt().to_string()).ok().unwrap();
                    }
                }

                // Whether the validation allows blank values
                map = map.map_put(atoms::allow_blank(), validation.get_allow_blank()).ok().unwrap();

                validations.push(map);
            }
        }

        Ok(validations)
    }));

    match result {
        Ok(Ok(validations)) => Ok(validations),
        Ok(Err(msg)) => Err((atoms::error(), msg)),
        Err(_) => Err((atoms::error(), "Error occurred in get_data_validations".to_string())),
    }
}

/// Get list validation rules
#[rustler::nif(name = "get_list_validations")]
pub fn get_list_validations_nif<'a>(
    env: rustler::Env<'a>,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_range: Option<String>,
) -> Result<Vec<rustler::Term<'a>>, (Atom, String)> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<Vec<rustler::Term<'a>>, String> {
        let spreadsheet_guard = resource.spreadsheet.lock().unwrap();
        let sheet = spreadsheet_guard
            .get_sheet_by_name(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        let mut list_validations = Vec::new();

        if let Some(data_validations) = sheet.get_data_validations() {
            for validation in data_validations.get_data_validation_list() {
                // Only include List validations
                if *validation.get_type() != DataValidationValues::List {
                    continue;
                }

                // If a specific cell range filter is provided
                if let Some(ref range_filter) = cell_range {
                    if validation.get_sequence_of_references().get_sqref() != *range_filter {
                        continue;
                    }
                }

                let mut map = rustler::types::map::map_new(env);

                // Common properties
                map = map.map_put(atoms::range(), validation.get_sequence_of_references().get_sqref().to_string()).ok().unwrap();
                map = map.map_put(atoms::rule_type(), atoms::list()).ok().unwrap();
                
                // Parse formula1 to get the list items
                let formula = validation.get_formula1().to_string();
                if formula.starts_with('"') && formula.ends_with('"') {
                    // Extract the comma-separated list
                    let list_str = formula[1..formula.len() - 1].to_string();
                    let list_items: Vec<String> = list_str.split(',').map(|s| s.to_string()).collect();
                    map = map.map_put(atoms::list_items(), list_items).ok().unwrap();
                } else {
                    // If formula doesn't start/end with quotes, it might be a range reference
                    map = map.map_put(atoms::formula1(), formula).ok().unwrap();
                }

                // Include error message details if present
                if *validation.get_show_error_message() {
                    map = map.map_put(atoms::show_error_message(), validation.get_show_error_message()).ok().unwrap();
                    if !validation.get_error_title().is_empty() {
                        map = map.map_put(atoms::error_title(), validation.get_error_title().to_string()).ok().unwrap();
                    }
                    if !validation.get_error_message().is_empty() {
                        map = map.map_put(atoms::error_message(), validation.get_error_message().to_string()).ok().unwrap();
                    }
                }

                // Include input message details if present
                if *validation.get_show_input_message() {
                    map = map.map_put(atoms::show_input_message(), validation.get_show_input_message()).ok().unwrap();
                    if !validation.get_prompt_title().is_empty() {
                        map = map.map_put(atoms::prompt_title(), validation.get_prompt_title().to_string()).ok().unwrap();
                    }
                    if !validation.get_prompt().is_empty() {
                        map = map.map_put(atoms::prompt_message(), validation.get_prompt().to_string()).ok().unwrap();
                    }
                }

                // Whether the validation allows blank values
                map = map.map_put(atoms::allow_blank(), validation.get_allow_blank()).ok().unwrap();

                list_validations.push(map);
            }
        }

        Ok(list_validations)
    }));

    match result {
        Ok(Ok(validations)) => Ok(validations),
        Ok(Err(msg)) => Err((atoms::error(), msg)),
        Err(_) => Err((atoms::error(), "Error occurred in get_list_validations".to_string())),
    }
}

/// Get decimal/number validation rules
#[rustler::nif(name = "get_number_validations")]
pub fn get_number_validations_nif<'a>(
    env: rustler::Env<'a>,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_range: Option<String>,
) -> Result<Vec<rustler::Term<'a>>, (Atom, String)> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<Vec<rustler::Term<'a>>, String> {
        let spreadsheet_guard = resource.spreadsheet.lock().unwrap();
        let sheet = spreadsheet_guard
            .get_sheet_by_name(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        let mut number_validations = Vec::new();

        if let Some(data_validations) = sheet.get_data_validations() {
            for validation in data_validations.get_data_validation_list() {
                // Only include Decimal and Whole validations
                if *validation.get_type() != DataValidationValues::Decimal && *validation.get_type() != DataValidationValues::Whole {
                    continue;
                }

                // If a specific cell range filter is provided
                if let Some(ref range_filter) = cell_range {
                    if validation.get_sequence_of_references().get_sqref() != *range_filter {
                        continue;
                    }
                }

                let mut map = rustler::types::map::map_new(env);

                // Common properties
                map = map.map_put(atoms::range(), validation.get_sequence_of_references().get_sqref().to_string()).ok().unwrap();
                let type_name = if *validation.get_type() == DataValidationValues::Decimal {"decimal"} else {"whole"};
                map = map.map_put(atoms::rule_type(), rustler::types::atom::Atom::from_str(env, type_name).unwrap()).ok().unwrap();
                
                // Operator
                let op_str = operator_to_string(validation.get_operator());
                map = map.map_put(atoms::operator(), rustler::types::atom::Atom::from_str(env, &op_str).unwrap()).ok().unwrap();

                // Formula values
                map = map.map_put(atoms::value1(), validation.get_formula1().to_string()).ok().unwrap();
                if !validation.get_formula2().is_empty() {
                    map = map.map_put(atoms::value2(), validation.get_formula2().to_string()).ok().unwrap();
                }

                // Include error message details if present
                if *validation.get_show_error_message() {
                    map = map.map_put(atoms::show_error_message(), validation.get_show_error_message()).ok().unwrap();
                    if !validation.get_error_title().is_empty() {
                        map = map.map_put(atoms::error_title(), validation.get_error_title().to_string()).ok().unwrap();
                    }
                    if !validation.get_error_message().is_empty() {
                        map = map.map_put(atoms::error_message(), validation.get_error_message().to_string()).ok().unwrap();
                    }
                }

                // Include input message details if present
                if *validation.get_show_input_message() {
                    map = map.map_put(atoms::show_input_message(), validation.get_show_input_message()).ok().unwrap();
                    if !validation.get_prompt_title().is_empty() {
                        map = map.map_put(atoms::prompt_title(), validation.get_prompt_title().to_string()).ok().unwrap();
                    }
                    if !validation.get_prompt().is_empty() {
                        map = map.map_put(atoms::prompt_message(), validation.get_prompt().to_string()).ok().unwrap();
                    }
                }

                // Whether the validation allows blank values
                map = map.map_put(atoms::allow_blank(), validation.get_allow_blank()).ok().unwrap();

                number_validations.push(map);
            }
        }

        Ok(number_validations)
    }));

    match result {
        Ok(Ok(validations)) => Ok(validations),
        Ok(Err(msg)) => Err((atoms::error(), msg)),
        Err(_) => Err((atoms::error(), "Error occurred in get_number_validations".to_string())),
    }
}

/// Get date validation rules
#[rustler::nif(name = "get_date_validations")]
pub fn get_date_validations_nif<'a>(
    env: rustler::Env<'a>,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_range: Option<String>,
) -> Result<Vec<rustler::Term<'a>>, (Atom, String)> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<Vec<rustler::Term<'a>>, String> {
        let spreadsheet_guard = resource.spreadsheet.lock().unwrap();
        let sheet = spreadsheet_guard
            .get_sheet_by_name(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        let mut date_validations = Vec::new();

        if let Some(data_validations) = sheet.get_data_validations() {
            for validation in data_validations.get_data_validation_list() {
                // Only include Date validations
                if *validation.get_type() != DataValidationValues::Date {
                    continue;
                }

                // If a specific cell range filter is provided
                if let Some(ref range_filter) = cell_range {
                    if validation.get_sequence_of_references().get_sqref() != *range_filter {
                        continue;
                    }
                }

                let mut map = rustler::types::map::map_new(env);

                // Common properties
                map = map.map_put(atoms::range(), validation.get_sequence_of_references().get_sqref().to_string()).ok().unwrap();
                map = map.map_put(atoms::rule_type(), atoms::date()).ok().unwrap();
                
                // Operator
                let op_str = operator_to_string(validation.get_operator());
                map = map.map_put(atoms::operator(), rustler::types::atom::Atom::from_str(env, &op_str).unwrap()).ok().unwrap();

                // Date values
                map = map.map_put(atoms::date1(), validation.get_formula1().to_string()).ok().unwrap();
                if !validation.get_formula2().is_empty() {
                    map = map.map_put(atoms::date2(), validation.get_formula2().to_string()).ok().unwrap();
                }

                // Include error message details if present
                if *validation.get_show_error_message() {
                    map = map.map_put(atoms::show_error_message(), validation.get_show_error_message()).ok().unwrap();
                    if !validation.get_error_title().is_empty() {
                        map = map.map_put(atoms::error_title(), validation.get_error_title().to_string()).ok().unwrap();
                    }
                    if !validation.get_error_message().is_empty() {
                        map = map.map_put(atoms::error_message(), validation.get_error_message().to_string()).ok().unwrap();
                    }
                }

                // Include input message details if present
                if *validation.get_show_input_message() {
                    map = map.map_put(atoms::show_input_message(), validation.get_show_input_message()).ok().unwrap();
                    if !validation.get_prompt_title().is_empty() {
                        map = map.map_put(atoms::prompt_title(), validation.get_prompt_title().to_string()).ok().unwrap();
                    }
                    if !validation.get_prompt().is_empty() {
                        map = map.map_put(atoms::prompt_message(), validation.get_prompt().to_string()).ok().unwrap();
                    }
                }

                // Whether the validation allows blank values
                map = map.map_put(atoms::allow_blank(), validation.get_allow_blank()).ok().unwrap();

                date_validations.push(map);
            }
        }

        Ok(date_validations)
    }));

    match result {
        Ok(Ok(validations)) => Ok(validations),
        Ok(Err(msg)) => Err((atoms::error(), msg)),
        Err(_) => Err((atoms::error(), "Error occurred in get_date_validations".to_string())),
    }
}

/// Get text length validation rules
#[rustler::nif(name = "get_text_length_validations")]
pub fn get_text_length_validations_nif<'a>(
    env: rustler::Env<'a>,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_range: Option<String>,
) -> Result<Vec<rustler::Term<'a>>, (Atom, String)> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<Vec<rustler::Term<'a>>, String> {
        let spreadsheet_guard = resource.spreadsheet.lock().unwrap();
        let sheet = spreadsheet_guard
            .get_sheet_by_name(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        let mut text_length_validations = Vec::new();

        if let Some(data_validations) = sheet.get_data_validations() {
            for validation in data_validations.get_data_validation_list() {
                // Only include TextLength validations
                if *validation.get_type() != DataValidationValues::TextLength {
                    continue;
                }

                // If a specific cell range filter is provided
                if let Some(ref range_filter) = cell_range {
                    if validation.get_sequence_of_references().get_sqref() != *range_filter {
                        continue;
                    }
                }

                let mut map = rustler::types::map::map_new(env);

                // Common properties
                map = map.map_put(atoms::range(), validation.get_sequence_of_references().get_sqref().to_string()).ok().unwrap();
                map = map.map_put(atoms::rule_type(), atoms::text_length()).ok().unwrap();
                
                // Operator
                let op_str = operator_to_string(validation.get_operator());
                map = map.map_put(atoms::operator(), rustler::types::atom::Atom::from_str(env, &op_str).unwrap()).ok().unwrap();

                // Length values
                // Try to convert formula values to integers
                let length1 = validation.get_formula1().parse::<i32>().unwrap_or(0);
                map = map.map_put(atoms::length1(), length1).ok().unwrap();

                if !validation.get_formula2().is_empty() {
                    let length2 = validation.get_formula2().parse::<i32>().unwrap_or(0);
                    map = map.map_put(atoms::length2(), length2).ok().unwrap();
                }

                // Include error message details if present
                if *validation.get_show_error_message() {
                    map = map.map_put(atoms::show_error_message(), validation.get_show_error_message()).ok().unwrap();
                    if !validation.get_error_title().is_empty() {
                        map = map.map_put(atoms::error_title(), validation.get_error_title().to_string()).ok().unwrap();
                    }
                    if !validation.get_error_message().is_empty() {
                        map = map.map_put(atoms::error_message(), validation.get_error_message().to_string()).ok().unwrap();
                    }
                }

                // Include input message details if present
                if *validation.get_show_input_message() {
                    map = map.map_put(atoms::show_input_message(), validation.get_show_input_message()).ok().unwrap();
                    if !validation.get_prompt_title().is_empty() {
                        map = map.map_put(atoms::prompt_title(), validation.get_prompt_title().to_string()).ok().unwrap();
                    }
                    if !validation.get_prompt().is_empty() {
                        map = map.map_put(atoms::prompt_message(), validation.get_prompt().to_string()).ok().unwrap();
                    }
                }

                // Whether the validation allows blank values
                map = map.map_put(atoms::allow_blank(), validation.get_allow_blank()).ok().unwrap();

                text_length_validations.push(map);
            }
        }

        Ok(text_length_validations)
    }));

    match result {
        Ok(Ok(validations)) => Ok(validations),
        Ok(Err(msg)) => Err((atoms::error(), msg)),
        Err(_) => Err((atoms::error(), "Error occurred in get_text_length_validations".to_string())),
    }
}

/// Get custom formula validation rules
#[rustler::nif(name = "get_custom_validations")]
pub fn get_custom_validations_nif<'a>(
    env: rustler::Env<'a>,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_range: Option<String>,
) -> Result<Vec<rustler::Term<'a>>, (Atom, String)> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<Vec<rustler::Term<'a>>, String> {
        let spreadsheet_guard = resource.spreadsheet.lock().unwrap();
        let sheet = spreadsheet_guard
            .get_sheet_by_name(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        let mut custom_validations = Vec::new();

        if let Some(data_validations) = sheet.get_data_validations() {
            for validation in data_validations.get_data_validation_list() {
                // Only include Custom validations
                if *validation.get_type() != DataValidationValues::Custom {
                    continue;
                }

                // If a specific cell range filter is provided
                if let Some(ref range_filter) = cell_range {
                    if validation.get_sequence_of_references().get_sqref() != *range_filter {
                        continue;
                    }
                }

                let mut map = rustler::types::map::map_new(env);

                // Common properties
                map = map.map_put(atoms::range(), validation.get_sequence_of_references().get_sqref().to_string()).ok().unwrap();
                map = map.map_put(atoms::rule_type(), atoms::custom()).ok().unwrap();
                
                // Formula
                map = map.map_put(atoms::formula(), validation.get_formula1().to_string()).ok().unwrap();

                // Include error message details if present
                if *validation.get_show_error_message() {
                    map = map.map_put(atoms::show_error_message(), validation.get_show_error_message()).ok().unwrap();
                    if !validation.get_error_title().is_empty() {
                        map = map.map_put(atoms::error_title(), validation.get_error_title().to_string()).ok().unwrap();
                    }
                    if !validation.get_error_message().is_empty() {
                        map = map.map_put(atoms::error_message(), validation.get_error_message().to_string()).ok().unwrap();
                    }
                }

                // Include input message details if present
                if *validation.get_show_input_message() {
                    map = map.map_put(atoms::show_input_message(), validation.get_show_input_message()).ok().unwrap();
                    if !validation.get_prompt_title().is_empty() {
                        map = map.map_put(atoms::prompt_title(), validation.get_prompt_title().to_string()).ok().unwrap();
                    }
                    if !validation.get_prompt().is_empty() {
                        map = map.map_put(atoms::prompt_message(), validation.get_prompt().to_string()).ok().unwrap();
                    }
                }

                // Whether the validation allows blank values
                map = map.map_put(atoms::allow_blank(), validation.get_allow_blank()).ok().unwrap();

                custom_validations.push(map);
            }
        }

        Ok(custom_validations)
    }));

    match result {
        Ok(Ok(validations)) => Ok(validations),
        Ok(Err(msg)) => Err((atoms::error(), msg)),
        Err(_) => Err((atoms::error(), "Error occurred in get_custom_validations".to_string())),
    }
}

/// Check if a sheet has any data validation rules
#[rustler::nif(name = "has_data_validations")]
pub fn has_data_validations_nif(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_range: Option<String>,
) -> Result<bool, (Atom, String)> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<bool, String> {
        let spreadsheet_guard = resource.spreadsheet.lock().unwrap();
        let sheet = spreadsheet_guard
            .get_sheet_by_name(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        if let Some(data_validations) = sheet.get_data_validations() {
            if data_validations.get_data_validation_list().is_empty() {
                return Ok(false);
            }

            if let Some(ref range_filter) = cell_range {
                for validation in data_validations.get_data_validation_list() {
                    if validation.get_sequence_of_references().get_sqref() == *range_filter {
                        return Ok(true);
                    }
                }
                // No validation found for the specified range
                return Ok(false);
            }

            // If we reach here, we have validations and no range filter
            Ok(true)
        } else {
            // No validations at all
            Ok(false)
        }
    }));

    match result {
        Ok(Ok(has_validations)) => Ok(has_validations),
        Ok(Err(msg)) => Err((atoms::error(), msg)),
        Err(_) => Err((atoms::error(), "Error occurred in has_data_validations".to_string())),
    }
}

/// Count data validation rules in a sheet
#[rustler::nif(name = "count_data_validations")]
pub fn count_data_validations_nif(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_range: Option<String>,
) -> Result<usize, (Atom, String)> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<usize, String> {
        let spreadsheet_guard = resource.spreadsheet.lock().unwrap();
        let sheet = spreadsheet_guard
            .get_sheet_by_name(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        if let Some(data_validations) = sheet.get_data_validations() {
            if let Some(ref range_filter) = cell_range {
                let mut count = 0;
                for validation in data_validations.get_data_validation_list() {
                    if validation.get_sequence_of_references().get_sqref() == *range_filter {
                        count += 1;
                    }
                }
                Ok(count)
            } else {
                Ok(data_validations.get_data_validation_list().len())
            }
        } else {
            Ok(0)
        }
    }));

    match result {
        Ok(Ok(count)) => Ok(count),
        Ok(Err(msg)) => Err((atoms::error(), msg)),
        Err(_) => Err((atoms::error(), "Error occurred in count_data_validations".to_string())),
    }
}
