use rustler::{Atom, Error as NifError, NifResult, ResourceArc};
use std::convert::TryInto;
use std::panic::{self, AssertUnwindSafe};
use umya_spreadsheet::{
    ConditionalFormatValueObject, ConditionalFormatValueObjectValues, ConditionalFormatValues,
    ConditionalFormatting, ConditionalFormattingRule, DataBar,
};

use crate::atoms;
use crate::helpers::style_helpers;
use crate::UmyaSpreadsheet;

#[rustler::nif]
fn add_data_bar(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_range: String,
    min_value: Option<(String, String)>,
    max_value: Option<(String, String)>,
    color: String,
) -> Result<Atom, (Atom, String)> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<Atom, String> {
        let mut spreadsheet = spreadsheet_resource
            .spreadsheet
            .lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        // Validate inputs
        if cell_range.trim().is_empty() {
            return Err("Cell range cannot be empty".to_string());
        }
        if color.trim().is_empty() {
            return Err("Color cannot be empty".to_string());
        }

        // Get worksheet by name
        let worksheet = spreadsheet
            .get_sheet_by_name_mut(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        // Create a new conditional formatting rule
        let mut rule = ConditionalFormattingRule::default();

        // Set the type to DataBar
        rule.set_type(ConditionalFormatValues::DataBar);

        // Create the data bar object
        let mut data_bar = DataBar::default();

        // Set min CFVO
        let mut cfvo_min = ConditionalFormatValueObject::default();
        if let Some((min_type, min_val)) = &min_value {
            match min_type.as_str() {
                "min" => cfvo_min.set_type(ConditionalFormatValueObjectValues::Min),
                "num" | "number" => {
                    cfvo_min.set_type(ConditionalFormatValueObjectValues::Number);
                    cfvo_min.set_val(min_val)
                }
                "percent" => {
                    cfvo_min.set_type(ConditionalFormatValueObjectValues::Percent);
                    cfvo_min.set_val(min_val)
                }
                "percentile" => {
                    cfvo_min.set_type(ConditionalFormatValueObjectValues::Percentile);
                    cfvo_min.set_val(min_val)
                }
                "formula" => {
                    cfvo_min.set_type(ConditionalFormatValueObjectValues::Formula);
                    cfvo_min.set_val(min_val)
                }
                _ => return Err(format!("Invalid min value type: {}", min_type)),
            };
        } else {
            cfvo_min.set_type(ConditionalFormatValueObjectValues::Min);
        }

        // Set max CFVO
        let mut cfvo_max = ConditionalFormatValueObject::default();
        if let Some((max_type, max_val)) = &max_value {
            match max_type.as_str() {
                "max" => cfvo_max.set_type(ConditionalFormatValueObjectValues::Max),
                "num" | "number" => {
                    cfvo_max.set_type(ConditionalFormatValueObjectValues::Number);
                    cfvo_max.set_val(max_val)
                }
                "percent" => {
                    cfvo_max.set_type(ConditionalFormatValueObjectValues::Percent);
                    cfvo_max.set_val(max_val)
                }
                "percentile" => {
                    cfvo_max.set_type(ConditionalFormatValueObjectValues::Percentile);
                    cfvo_max.set_val(max_val)
                }
                "formula" => {
                    cfvo_max.set_type(ConditionalFormatValueObjectValues::Formula);
                    cfvo_max.set_val(max_val)
                }
                _ => return Err(format!("Invalid max value type: {}", max_type)),
            };
        } else {
            cfvo_max.set_type(ConditionalFormatValueObjectValues::Max);
        }

        // Add CFVOs to the data bar
        data_bar.add_cfvo_collection(cfvo_min);
        data_bar.add_cfvo_collection(cfvo_max);

        // Set the color
        let bar_color = style_helpers::parse_color(&color);
        data_bar.add_color_collection(bar_color);

        // Set the data bar in the rule
        rule.set_data_bar(data_bar);

        // Add the rule to conditional formatting
        let mut conditional_formatting = ConditionalFormatting::default();
        let mut sequence = umya_spreadsheet::SequenceOfReferences::default();
        sequence.set_sqref(cell_range);
        conditional_formatting.set_sequence_of_references(sequence);
        conditional_formatting.add_conditional_collection(rule);

        // Add the conditional formatting to the worksheet
        worksheet.add_conditional_formatting_collection(conditional_formatting);

        Ok(atoms::ok())
    }));

    match result {
        Ok(Ok(atom)) => Ok(atom),
        Ok(Err(err_msg)) => Err((atoms::error(), err_msg)),
        Err(_) => Err((
            atoms::error(),
            "Error occurred in add_data_bar operation".to_string(),
        )),
    }
}

#[rustler::nif]
fn add_top_bottom_rule(
    spreadsheet_resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_range: String,
    rule_type: String,
    rank: i32, // Keep as i32 for API compatibility, but convert to u32 internally if valid
    percent: bool,
    format_style: String,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get worksheet by name
        if let Some(worksheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) {
            // Create a new conditional formatting rule
            let mut rule = ConditionalFormattingRule::default();

            // Validate and convert rank to u32
            let rank_u32: u32 = if rank < 0 {
                return Err("Rank must be a positive number".to_string());
            } else {
                rank.try_into()
                    .map_err(|_| "Invalid rank value".to_string())?
            };

            // Set the rule type based on top/bottom and percent
            match rule_type.to_lowercase().as_str() {
                "top" => {
                    if percent {
                        rule.set_type(ConditionalFormatValues::Top10);
                        rule.set_percent(true);
                    } else {
                        rule.set_type(ConditionalFormatValues::Top10);
                        rule.set_percent(false);
                    }
                    rule.set_rank(rank_u32);
                    rule.set_bottom(false);
                }
                "bottom" => {
                    if percent {
                        rule.set_type(ConditionalFormatValues::Top10);
                        rule.set_percent(true);
                    } else {
                        rule.set_type(ConditionalFormatValues::Top10);
                        rule.set_percent(false);
                    }
                    rule.set_rank(rank_u32);
                    rule.set_bottom(true);
                }
                _ => return Err(format!("Invalid rule type: {}", rule_type)),
            }

            // Set style if provided
            if let Some(fill) = style_helpers::create_fill_style(&format_style) {
                rule.set_style(fill);
            }

            // Add the rule to conditional formatting
            let mut conditional_formatting = ConditionalFormatting::default();
            let mut sequence = umya_spreadsheet::SequenceOfReferences::default();
            sequence.set_sqref(cell_range);
            conditional_formatting.set_sequence_of_references(sequence);
            conditional_formatting.add_conditional_collection(rule);

            // Add the conditional formatting to the worksheet
            worksheet.add_conditional_formatting_collection(conditional_formatting);

            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => Err(NifError::Term(Box::new((atoms::error(), msg)))),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Error occurred in add_top_bottom_rule".to_string(),
        )))),
    }
}

#[rustler::nif]
fn add_text_rule(
    spreadsheet_resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_range: String,
    rule_type: String,
    text: String,
    format_style: String,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get worksheet by name
        if let Some(worksheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) {
            // Create a new conditional formatting rule
            let mut rule = ConditionalFormattingRule::default();

            // Set the text and type based on the rule type
            match rule_type.to_lowercase().as_str() {
                "contains" => {
                    rule.set_type(ConditionalFormatValues::ContainsText);
                    rule.set_text(&text);
                }
                "beginswith" | "begins_with" => {
                    rule.set_type(ConditionalFormatValues::BeginsWith);
                    rule.set_text(&text);
                }
                "endswith" | "ends_with" => {
                    rule.set_type(ConditionalFormatValues::EndsWith);
                    rule.set_text(&text);
                }
                _ => return Err(format!("Invalid text rule type: {}", rule_type)),
            }

            // Set style if provided
            if let Some(fill) = style_helpers::create_fill_style(&format_style) {
                rule.set_style(fill);
            }

            // Add the rule to conditional formatting
            let mut conditional_formatting = ConditionalFormatting::default();
            let mut sequence = umya_spreadsheet::SequenceOfReferences::default();
            sequence.set_sqref(cell_range);
            conditional_formatting.set_sequence_of_references(sequence);
            conditional_formatting.add_conditional_collection(rule);

            // Add the conditional formatting to the worksheet
            worksheet.add_conditional_formatting_collection(conditional_formatting);

            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => Err(NifError::Term(Box::new((atoms::error(), msg)))),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Error occurred in add_text_rule".to_string(),
        )))),
    }
}

// Note: add_color_scale is now implemented in conditional_formatting.rs
// This implementation was removed to prevent duplicate function errors
