use rustler::ResourceArc;
use umya_spreadsheet::{
    ColorScale,
    ConditionalFormatValueObject,
    ConditionalFormatValueObjectValues,
    ConditionalFormatValues,
    ConditionalFormattingOperatorValues, // Removed ConditionalFormatting
    ConditionalFormattingRule,
    Formula,
    Style,
};

// Removed unused imports: Env, Term

// Use the public modules from lib.rs
use crate::custom_structs::{CustomColor, CustomFont};
use crate::UmyaSpreadsheet;

// Helper function to parse font_json into umya_spreadsheet::Font
fn parse_font(font_json: Option<CustomFont>) -> Option<umya_spreadsheet::Font> {
    font_json.map(|cf| {
        let mut font = umya_spreadsheet::Font::default();
        if let Some(b) = cf.bold {
            font.set_bold(b);
        }
        if let Some(i) = cf.italic {
            font.set_italic(i);
        }
        if let Some(c) = cf.color {
            let mut color = umya_spreadsheet::Color::default();
            color.set_argb(c.argb);
            font.set_color(color);
        }
        font
    })
}

// Helper function to convert string type to ConditionalFormatValueObjectValues
fn parse_cfvo_type(type_str: String) -> Result<ConditionalFormatValueObjectValues, String> {
    if type_str.is_empty() {
        return Ok(ConditionalFormatValueObjectValues::Min);
    }
    
    match type_str.to_lowercase().as_str() {
        "formula" => Ok(ConditionalFormatValueObjectValues::Formula),
        "max" => Ok(ConditionalFormatValueObjectValues::Max),
        "min" => Ok(ConditionalFormatValueObjectValues::Min),
        "num" | "number" => Ok(ConditionalFormatValueObjectValues::Number),
        "percent" => Ok(ConditionalFormatValueObjectValues::Percent),
        "percentile" => Ok(ConditionalFormatValueObjectValues::Percentile),
        _ => Err(format!("Unsupported CFVO type: {}", type_str)),
    }
}

#[rustler::nif]
pub fn add_cell_value_rule(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    range: String,
    operator: String,
    value1: String,
    value2: Option<String>,
    format_style: String,
) -> Result<bool, String> {
    let mut spreadsheet_guard = resource.spreadsheet.lock().unwrap();
    let sheet = spreadsheet_guard
        .get_sheet_by_name_mut(&sheet_name)
        .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

    let mut rule = ConditionalFormattingRule::default();
    rule.set_type(ConditionalFormatValues::CellIs);

    // Set operator
    let operator_str = operator.as_str();
    let op_val = match operator_str {
        "equal" => ConditionalFormattingOperatorValues::Equal,
        "not_equal" => ConditionalFormattingOperatorValues::NotEqual,
        "greater_than" => ConditionalFormattingOperatorValues::GreaterThan,
        "greater_than_or_equal" => ConditionalFormattingOperatorValues::GreaterThanOrEqual,
        "less_than" => ConditionalFormattingOperatorValues::LessThan,
        "less_than_or_equal" => ConditionalFormattingOperatorValues::LessThanOrEqual,
        "between" => ConditionalFormattingOperatorValues::Between,
        "not_between" => ConditionalFormattingOperatorValues::NotBetween,
        "containing" | "containstext" => ConditionalFormattingOperatorValues::ContainsText,
        "not_containing" | "notcontains" => ConditionalFormattingOperatorValues::NotContains,
        "beginning_with" | "beginswith" => ConditionalFormattingOperatorValues::BeginsWith,
        "ending_with" | "endswith" => ConditionalFormattingOperatorValues::EndsWith,
        _ => return Err(format!("Unsupported operator: {}", operator)),
    };
    rule.set_operator(op_val);

    // Handle formula value(s) based on operator
    if operator.as_str() == "between" || operator.as_str() == "not_between" {
        // Two formulas needed for between/not between
        if let Some(val2) = &value2 {
            // First value
            let mut formula1 = Formula::default();
            formula1.set_string_value(value1.clone());
            rule.set_formula(formula1);
            
            // For the second value in between operators
            // Just set the formula again with the second value
            // Since setting formula replaces the previous one
            let mut formula2 = Formula::default();
            formula2.set_string_value(val2.clone());
            rule.set_formula(formula2);
        } else {
            return Err("Second value is required for between/not between operators".to_string());
        }
    } else {
        // Single formula for other operators
        let mut formula = Formula::default();
        formula.set_string_value(value1);
        rule.set_formula(formula);
    }

    // Create a style from the format_style string (expected to be a color code)
    let mut style_to_apply = Style::default();
    let mut style_modified = false;

    // If format_style is a color code, create a fill with that color
    if !format_style.is_empty() {
        let mut fill = umya_spreadsheet::Fill::default();
        let mut pattern_fill = umya_spreadsheet::PatternFill::default();
        pattern_fill.set_pattern_type(umya_spreadsheet::PatternValues::Solid);

        // Create a color from the format_style string
        let mut fg_color_obj = umya_spreadsheet::Color::default();
        // Handle different formats of color strings
        if format_style.starts_with('#') && format_style.len() == 7 {
            // Format: #RRGGBB - Convert to FFRRGGBB for Excel
            fg_color_obj.set_argb(&format!("FF{}", &format_style[1..]));
        } else {
            // Already in ARGB format or other format
            fg_color_obj.set_argb(&format_style);
        }
        
        pattern_fill.set_foreground_color(fg_color_obj);
        fill.set_pattern_fill(pattern_fill);
        style_to_apply.set_fill(fill);
        style_modified = true;
    }

    if style_modified {
        rule.set_style(style_to_apply);
    }

    let mut conditional_formatting_collection = umya_spreadsheet::ConditionalFormatting::default();
    conditional_formatting_collection
        .get_sequence_of_references_mut()
        .set_sqref(range); // Set sqref here
    conditional_formatting_collection.add_conditional_collection(rule);
    sheet.add_conditional_formatting_collection(conditional_formatting_collection);

    Ok(true)
}

#[rustler::nif]
pub fn add_color_scale(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    range: String,
    min_type: String,
    min_value: Option<String>,
    min_color: Option<CustomColor>,
    mid_type: Option<String>,
    mid_value: Option<String>,
    mid_color: Option<CustomColor>,
    max_type: String,
    max_value: Option<String>,
    max_color: Option<CustomColor>,
) -> Result<bool, String> {
    let mut spreadsheet_guard = resource.spreadsheet.lock().unwrap();
    let sheet = spreadsheet_guard
        .get_sheet_by_name_mut(&sheet_name)
        .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

    let mut rule = ConditionalFormattingRule::default();
    rule.set_type(ConditionalFormatValues::ColorScale);

    let mut color_scale_value = ColorScale::default();

    // Min
    let mut min_cfvo = ConditionalFormatValueObject::default();
    if let Ok(cfvo_type) = parse_cfvo_type(min_type.clone()) {
        min_cfvo.set_type(cfvo_type);
    } else {
        min_cfvo.set_type(ConditionalFormatValueObjectValues::Min);
    }
    
    if let Some(val) = min_value {
        if !val.is_empty() {
            min_cfvo.set_val(val);
        }
    }
    color_scale_value.add_cfvo_collection(min_cfvo);

    let mut min_color_obj = umya_spreadsheet::Color::default();
    if let Some(color) = min_color {
        min_color_obj.set_argb(&color.argb);
    } else {
        min_color_obj.set_argb("FFFFFFFF"); // Default to white
    }
    color_scale_value.add_color_collection(min_color_obj);

    // Mid (optional)
    if let Some(mt_str) = mid_type {
        let mut mid_cfvo = ConditionalFormatValueObject::default();
        if let Ok(cfvo_type) = parse_cfvo_type(mt_str) {
            mid_cfvo.set_type(cfvo_type);
        } else {
            mid_cfvo.set_type(ConditionalFormatValueObjectValues::Percentile);
            mid_cfvo.set_val("50"); // Default to 50th percentile
        }
        
        if let Some(val) = mid_value {
            if !val.is_empty() {
                mid_cfvo.set_val(val);
            }
        }
        color_scale_value.add_cfvo_collection(mid_cfvo);

        let mut mid_color_obj = umya_spreadsheet::Color::default();
        if let Some(color) = &mid_color {
            mid_color_obj.set_argb(&color.argb);
        } else {
            mid_color_obj.set_argb("FFFFFF00"); // Default to yellow
        }
        color_scale_value.add_color_collection(mid_color_obj);
    }

    // Max
    let mut max_cfvo = ConditionalFormatValueObject::default();
    if let Ok(cfvo_type) = parse_cfvo_type(max_type.clone()) {
        max_cfvo.set_type(cfvo_type);
    } else {
        max_cfvo.set_type(ConditionalFormatValueObjectValues::Max);
    }
    
    if let Some(val) = max_value {
        if !val.is_empty() {
            max_cfvo.set_val(val);
        }
    }
    color_scale_value.add_cfvo_collection(max_cfvo);

    let mut max_color_obj = umya_spreadsheet::Color::default();
    if let Some(color) = max_color {
        max_color_obj.set_argb(&color.argb);
    } else {
        max_color_obj.set_argb("FF00FF00"); // Default to green
    }
    color_scale_value.add_color_collection(max_color_obj);

    rule.set_color_scale(color_scale_value);

    let mut conditional_formatting_collection = umya_spreadsheet::ConditionalFormatting::default();
    conditional_formatting_collection
        .get_sequence_of_references_mut()
        .set_sqref(range); // Set sqref here
    conditional_formatting_collection.add_conditional_collection(rule);
    sheet.add_conditional_formatting_collection(conditional_formatting_collection); // Pass only one argument

    Ok(true)
}
