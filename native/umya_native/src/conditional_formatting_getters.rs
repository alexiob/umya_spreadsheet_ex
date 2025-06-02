use rustler::{Atom, Encoder, ResourceArc};
use std::panic::{self, AssertUnwindSafe};
use umya_spreadsheet::{
    ConditionalFormatValueObjectValues, ConditionalFormatValues,
    ConditionalFormattingOperatorValues, ConditionalFormattingRule,
};

use crate::atoms;
use crate::custom_structs::CustomColor;
use crate::UmyaSpreadsheet;

// Helper function to convert ConditionalFormatValueObjectValues to string
fn cfvo_type_to_string(value: &ConditionalFormatValueObjectValues) -> String {
    match value {
        ConditionalFormatValueObjectValues::Formula => "formula".to_string(),
        ConditionalFormatValueObjectValues::Max => "max".to_string(),
        ConditionalFormatValueObjectValues::Min => "min".to_string(),
        ConditionalFormatValueObjectValues::Number => "number".to_string(),
        ConditionalFormatValueObjectValues::Percent => "percent".to_string(),
        ConditionalFormatValueObjectValues::Percentile => "percentile".to_string(),
    }
}

// Helper function to convert ConditionalFormattingOperatorValues to string
fn operator_to_string(op: &ConditionalFormattingOperatorValues) -> String {
    match op {
        ConditionalFormattingOperatorValues::Equal => "equal".to_string(),
        ConditionalFormattingOperatorValues::NotEqual => "not_equal".to_string(),
        ConditionalFormattingOperatorValues::GreaterThan => "greater_than".to_string(),
        ConditionalFormattingOperatorValues::LessThan => "less_than".to_string(),
        ConditionalFormattingOperatorValues::GreaterThanOrEqual => {
            "greater_than_or_equal".to_string()
        }
        ConditionalFormattingOperatorValues::LessThanOrEqual => "less_than_or_equal".to_string(),
        ConditionalFormattingOperatorValues::Between => "between".to_string(),
        ConditionalFormattingOperatorValues::NotBetween => "not_between".to_string(),
        ConditionalFormattingOperatorValues::ContainsText => "contains".to_string(),
        ConditionalFormattingOperatorValues::NotContains => "not_contains".to_string(),
        ConditionalFormattingOperatorValues::BeginsWith => "begins_with".to_string(),
        ConditionalFormattingOperatorValues::EndsWith => "ends_with".to_string(),
    }
}

// Helper function to convert ConditionalFormatValues to string
fn format_type_to_string(format_type: &ConditionalFormatValues) -> String {
    match format_type {
        ConditionalFormatValues::CellIs => "cell_is".to_string(),
        ConditionalFormatValues::ColorScale => "color_scale".to_string(),
        ConditionalFormatValues::DataBar => "data_bar".to_string(),
        ConditionalFormatValues::IconSet => "icon_set".to_string(),
        ConditionalFormatValues::Top10 => "top10".to_string(),
        ConditionalFormatValues::AboveAverage => "above_average".to_string(),
        ConditionalFormatValues::BeginsWith => "begins_with".to_string(),
        ConditionalFormatValues::ContainsText => "contains_text".to_string(),
        ConditionalFormatValues::EndsWith => "ends_with".to_string(),
        _ => "unknown".to_string(),
    }
}

// Helper to convert a color to a CustomColor
fn color_to_custom_color(color: &umya_spreadsheet::Color) -> CustomColor {
    CustomColor {
        argb: color.get_argb().to_string(),
    }
}

// Helper function to extract cell value rule details
fn extract_cell_value_rule<'a>(
    env: rustler::Env<'a>,
    rule: &ConditionalFormattingRule,
) -> rustler::Term<'a> {
    let operator = operator_to_string(rule.get_operator());

    // Get formula values
    let formula = match rule.get_formula() {
        Some(formula) => formula.get_address_str(),
        None => "".to_string(),
    };

    // Extract style information (focusing on cell color)
    let format_style = match rule.get_style() {
        Some(style) => {
            if let Some(fill) = style.get_fill() {
                if let Some(pattern_fill) = fill.get_pattern_fill() {
                    if let Some(fg_color) = pattern_fill.get_foreground_color() {
                        fg_color.get_argb().to_string()
                    } else {
                        "".to_string()
                    }
                } else {
                    "".to_string()
                }
            } else {
                "".to_string()
            }
        }
        None => "".to_string(),
    };

    let mut map = rustler::types::map::map_new(env);
    map = map
        .map_put(atoms::rule_type(), atoms::cell_is())
        .ok()
        .unwrap();
    map = map
        .map_put(
            atoms::operator(),
            rustler::types::atom::Atom::from_str(env, operator.as_str()).unwrap(),
        )
        .ok()
        .unwrap();
    map = map.map_put(atoms::formula(), formula.clone()).ok().unwrap();
    map = map
        .map_put(atoms::format_style(), format_style)
        .ok()
        .unwrap();

    map
}

// Helper function to extract color scale rule details
fn extract_color_scale<'a>(
    env: rustler::Env<'a>,
    rule: &ConditionalFormattingRule,
) -> rustler::Term<'a> {
    match rule.get_color_scale() {
        Some(color_scale) => {
            let cfvos = color_scale.get_cfvo_collection();
            let colors = color_scale.get_color_collection();

            let mut map = rustler::types::map::map_new(env);
            map = map
                .map_put(atoms::rule_type(), atoms::color_scale())
                .ok()
                .unwrap();

            // Extract min, mid (if present), and max values
            if cfvos.len() >= 1 && colors.len() >= 1 {
                let min_cfvo = &cfvos[0];
                let min_color = &colors[0];

                let min_type = cfvo_type_to_string(min_cfvo.get_type());
                let min_value = min_cfvo.get_val().to_string();
                let min_color_map = color_to_custom_color(min_color);

                map = map
                    .map_put(
                        atoms::min_type(),
                        rustler::types::atom::Atom::from_str(env, min_type.as_str()).unwrap(),
                    )
                    .ok()
                    .unwrap();
                map = map.map_put(atoms::min_value(), min_value).ok().unwrap();
                map = map.map_put(atoms::min_color(), min_color_map).ok().unwrap();
            }

            // Mid values if available (3-color scale)
            if cfvos.len() >= 2 && colors.len() >= 2 && cfvos.len() == 3 && colors.len() == 3 {
                let mid_cfvo = &cfvos[1];
                let mid_color = &colors[1];

                let mid_type = cfvo_type_to_string(mid_cfvo.get_type());
                let mid_value = mid_cfvo.get_val().to_string();
                let mid_color_map = color_to_custom_color(mid_color);

                map = map
                    .map_put(
                        atoms::mid_type(),
                        rustler::types::atom::Atom::from_str(env, mid_type.as_str()).unwrap(),
                    )
                    .ok()
                    .unwrap();
                map = map.map_put(atoms::mid_value(), mid_value).ok().unwrap();
                map = map.map_put(atoms::mid_color(), mid_color_map).ok().unwrap();
            }

            // Max values (last in the collections)
            if cfvos.len() >= 2 && colors.len() >= 2 {
                let max_index = cfvos.len() - 1;
                let max_cfvo = &cfvos[max_index];
                let max_color = &colors[max_index];

                let max_type = cfvo_type_to_string(max_cfvo.get_type());
                let max_value = max_cfvo.get_val().to_string();
                let max_color_map = color_to_custom_color(max_color);

                map = map
                    .map_put(
                        atoms::max_type(),
                        rustler::types::atom::Atom::from_str(env, max_type.as_str()).unwrap(),
                    )
                    .ok()
                    .unwrap();
                map = map.map_put(atoms::max_value(), max_value).ok().unwrap();
                map = map.map_put(atoms::max_color(), max_color_map).ok().unwrap();
            }

            map
        }
        None => rustler::types::map::map_new(env),
    }
}

// Helper function to extract data bar rule details
fn extract_data_bar<'a>(
    env: rustler::Env<'a>,
    rule: &ConditionalFormattingRule,
) -> rustler::Term<'a> {
    match rule.get_data_bar() {
        Some(data_bar) => {
            let cfvos = data_bar.get_cfvo_collection();
            let colors = data_bar.get_color_collection();

            let mut map = rustler::types::map::map_new(env);
            map = map
                .map_put(atoms::rule_type(), atoms::data_bar())
                .ok()
                .unwrap();

            if cfvos.len() >= 2 && colors.len() >= 1 {
                let min_cfvo = &cfvos[0];
                let max_cfvo = &cfvos[1];
                let color = &colors[0];

                // Extract min and max values
                let min_type = cfvo_type_to_string(min_cfvo.get_type());
                let min_value = min_cfvo.get_val().to_string();
                let max_type = cfvo_type_to_string(max_cfvo.get_type());
                let max_value = max_cfvo.get_val().to_string();

                // Create optional tuples for min and max if they're not the default types
                let min_tuple = if min_type == "min" && min_value.is_empty() {
                    // Convert the atom to a Term to make the types compatible
                    let atom = rustler::types::atom::Atom::from_str(env, "nil").unwrap();
                    atom.encode(env)
                } else {
                    let min_type_atom =
                        rustler::types::atom::Atom::from_str(env, min_type.as_str()).unwrap();
                    let inner_tuple = rustler::types::tuple::make_tuple(
                        env,
                        &[min_type_atom.encode(env), min_value.clone().encode(env)],
                    );
                    inner_tuple
                };

                let max_tuple = if max_type == "max" && max_value.is_empty() {
                    // Convert the atom to a Term to make the types compatible
                    let atom = rustler::types::atom::Atom::from_str(env, "nil").unwrap();
                    atom.encode(env)
                } else {
                    let max_type_atom =
                        rustler::types::atom::Atom::from_str(env, max_type.as_str()).unwrap();
                    let inner_tuple = rustler::types::tuple::make_tuple(
                        env,
                        &[max_type_atom.encode(env), max_value.clone().encode(env)],
                    );
                    inner_tuple
                };

                // Get color
                let color_value = color.get_argb().to_string();

                map = map.map_put(atoms::min_value(), min_tuple).ok().unwrap();
                map = map.map_put(atoms::max_value(), max_tuple).ok().unwrap();
                map = map.map_put(atoms::color(), color_value).ok().unwrap();
            }

            map
        }
        None => rustler::types::map::map_new(env),
    }
}

// Helper function to extract icon set rule details
fn extract_icon_set<'a>(
    env: rustler::Env<'a>,
    rule: &ConditionalFormattingRule,
) -> rustler::Term<'a> {
    match rule.get_icon_set() {
        Some(icon_set) => {
            let cfvos = icon_set.get_cfvo_collection();
            // Since IconSetValues is not available, determine the type by the number of items
            let icon_set_type = match cfvos.len() {
                3 => "3arrows", // Default to 3 arrows when there are 3 items
                4 => "4arrows", // Default to 4 arrows when there are 4 items
                5 => "5arrows", // Default to 5 arrows when there are 5 items
                _ => "unknown",
            }
            .to_string();

            let mut map = rustler::types::map::map_new(env);
            map = map
                .map_put(atoms::rule_type(), atoms::icon_set())
                .ok()
                .unwrap();
            map = map
                .map_put(atoms::icon_style(), icon_set_type)
                .ok()
                .unwrap();

            // Extract thresholds
            let mut thresholds = Vec::new();
            for cfvo in cfvos {
                let threshold_type = cfvo_type_to_string(cfvo.get_type());
                let threshold_value = cfvo.get_val().to_string();

                thresholds.push((threshold_type, threshold_value));
            }

            map = map.map_put(atoms::thresholds(), thresholds).ok().unwrap();

            map
        }
        None => rustler::types::map::map_new(env),
    }
}

// Helper function to extract top/bottom rules
fn extract_top_bottom_rule<'a>(
    env: rustler::Env<'a>,
    rule: &ConditionalFormattingRule,
) -> rustler::Term<'a> {
    let mut map = rustler::types::map::map_new(env);
    map = map
        .map_put(atoms::rule_type(), atoms::top_bottom())
        .ok()
        .unwrap();

    // Get rank value
    let rank = *rule.get_rank();

    // Determine rule type (top or bottom)
    let rule_type = if *rule.get_bottom() { "bottom" } else { "top" };

    // Check if percent is set
    let percent = *rule.get_percent();

    // Extract style information (focusing on cell color)
    let format_style = match rule.get_style() {
        Some(style) => {
            if let Some(fill) = style.get_fill() {
                if let Some(pattern_fill) = fill.get_pattern_fill() {
                    if let Some(fg_color) = pattern_fill.get_foreground_color() {
                        fg_color.get_argb().to_string()
                    } else {
                        "".to_string()
                    }
                } else {
                    "".to_string()
                }
            } else {
                "".to_string()
            }
        }
        None => "".to_string(),
    };

    map = map
        .map_put(
            atoms::rule_type_value(),
            rustler::types::atom::Atom::from_str(env, rule_type).unwrap(),
        )
        .ok()
        .unwrap();
    map = map.map_put(atoms::rank(), rank).ok().unwrap();
    map = map.map_put(atoms::percent(), percent).ok().unwrap();
    map = map
        .map_put(atoms::format_style(), format_style)
        .ok()
        .unwrap();

    map
}

// Helper function to extract above/below average rules
fn extract_above_below_average_rule<'a>(
    env: rustler::Env<'a>,
    rule: &ConditionalFormattingRule,
) -> rustler::Term<'a> {
    let mut map = rustler::types::map::map_new(env);
    map = map
        .map_put(atoms::rule_type(), atoms::above_below_average())
        .ok()
        .unwrap();

    // Determine rule type
    let rule_type = if *rule.get_above_average() {
        if *rule.get_equal_average() {
            "above_equal"
        } else {
            "above"
        }
    } else {
        if *rule.get_equal_average() {
            "below_equal"
        } else {
            "below"
        }
    };

    // Get standard deviation
    let std_dev = *rule.get_std_dev();

    // Extract style information
    let format_style = match rule.get_style() {
        Some(style) => {
            if let Some(fill) = style.get_fill() {
                if let Some(pattern_fill) = fill.get_pattern_fill() {
                    if let Some(fg_color) = pattern_fill.get_foreground_color() {
                        fg_color.get_argb().to_string()
                    } else {
                        "".to_string()
                    }
                } else {
                    "".to_string()
                }
            } else {
                "".to_string()
            }
        }
        None => "".to_string(),
    };

    map = map
        .map_put(
            atoms::rule_type_value(),
            rustler::types::atom::Atom::from_str(env, rule_type).unwrap(),
        )
        .ok()
        .unwrap();
    map = map.map_put(atoms::std_dev(), std_dev).ok().unwrap();
    map = map
        .map_put(atoms::format_style(), format_style)
        .ok()
        .unwrap();

    map
}

// Helper function to extract text rules
fn extract_text_rule<'a>(
    env: rustler::Env<'a>,
    rule: &ConditionalFormattingRule,
) -> rustler::Term<'a> {
    let mut map = rustler::types::map::map_new(env);

    // Determine the specific text rule type
    let rule_type = format_type_to_string(rule.get_type());
    let operator = if rule_type == "contains_text" {
        "contains"
    } else if rule_type == "begins_with" {
        "begins_with"
    } else if rule_type == "ends_with" {
        "ends_with"
    } else {
        "unknown"
    };

    map = map
        .map_put(atoms::rule_type(), atoms::text_rule())
        .ok()
        .unwrap();
    map = map
        .map_put(
            atoms::operator(),
            rustler::types::atom::Atom::from_str(env, operator).unwrap(),
        )
        .ok()
        .unwrap();
    map = map
        .map_put(atoms::text(), rule.get_text().to_string())
        .ok()
        .unwrap();

    // Extract style information
    let format_style = match rule.get_style() {
        Some(style) => {
            if let Some(fill) = style.get_fill() {
                if let Some(pattern_fill) = fill.get_pattern_fill() {
                    if let Some(fg_color) = pattern_fill.get_foreground_color() {
                        fg_color.get_argb().to_string()
                    } else {
                        "".to_string()
                    }
                } else {
                    "".to_string()
                }
            } else {
                "".to_string()
            }
        }
        None => "".to_string(),
    };

    map = map
        .map_put(atoms::format_style(), format_style)
        .ok()
        .unwrap();

    map
}

#[rustler::nif(name = "get_conditional_formatting_rules")]
pub fn get_conditional_formatting_rules_nif<'a>(
    env: rustler::Env<'a>,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    range: Option<String>,
) -> Result<Vec<rustler::Term<'a>>, (Atom, String)> {
    let result = panic::catch_unwind(AssertUnwindSafe(
        || -> Result<Vec<rustler::Term<'a>>, String> {
            let spreadsheet_guard = resource.spreadsheet.lock().unwrap();
            let sheet = spreadsheet_guard
                .get_sheet_by_name(&sheet_name)
                .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

            let cf_collections = sheet.get_conditional_formatting_collection();
            let mut rules = Vec::new();

            for cf in cf_collections {
                let sqref = cf.get_sequence_of_references().get_sqref();

                // If range filter is specified, only include rules that match
                if let Some(ref filter_range) = range {
                    if sqref != *filter_range {
                        continue;
                    }
                }

                for rule in cf.get_conditional_collection() {
                    // Create a base map with common properties
                    let mut base_map = rustler::types::map::map_new(env);
                    base_map = base_map
                        .map_put(atoms::range(), sqref.clone())
                        .ok()
                        .unwrap();

                    // Create a rule-specific map based on the type
                    let rule_type = rule.get_type();
                    let rule_map = match rule_type {
                        ConditionalFormatValues::CellIs => extract_cell_value_rule(env, rule),
                        ConditionalFormatValues::ColorScale => extract_color_scale(env, rule),
                        ConditionalFormatValues::DataBar => extract_data_bar(env, rule),
                        ConditionalFormatValues::IconSet => extract_icon_set(env, rule),
                        ConditionalFormatValues::Top10 => extract_top_bottom_rule(env, rule),
                        ConditionalFormatValues::AboveAverage => {
                            extract_above_below_average_rule(env, rule)
                        }
                        ConditionalFormatValues::ContainsText
                        | ConditionalFormatValues::BeginsWith
                        | ConditionalFormatValues::EndsWith => extract_text_rule(env, rule),
                        _ => rustler::types::map::map_new(env),
                    };

                    // Since map_iter isn't available, we'll need to merge based on known keys
                    let mut merged_map = base_map;

                    // Add all fixed keys we expect in our maps
                    // This isn't elegant but will get our code working
                    let keys = [
                        atoms::rule_type(),
                        atoms::operator(),
                        atoms::formula(),
                        atoms::format_style(),
                        atoms::min_type(),
                        atoms::min_value(),
                        atoms::min_color(),
                        atoms::mid_type(),
                        atoms::mid_value(),
                        atoms::mid_color(),
                        atoms::max_type(),
                        atoms::max_value(),
                        atoms::max_color(),
                        atoms::color(),
                        atoms::icon_style(),
                        atoms::thresholds(),
                        atoms::rank(),
                        atoms::percent(),
                        atoms::std_dev(),
                        atoms::text(),
                        atoms::rule_type_value(),
                    ];

                    // Try to get values for each key from rule_map
                    for key in &keys {
                        if let Ok(value) = rule_map.map_get(*key) {
                            merged_map = merged_map.map_put(*key, value).ok().unwrap();
                        }
                    }

                    rules.push(merged_map);
                }
            }

            Ok(rules)
        },
    ));

    match result {
        Ok(Ok(rules)) => Ok(rules),
        Ok(Err(msg)) => Err((atoms::error(), msg)),
        Err(_) => Err((
            atoms::error(),
            "Error occurred in get_conditional_formatting_rules".to_string(),
        )),
    }
}

// Function to get specific types of conditional formatting rules
#[rustler::nif(name = "get_cell_value_rules")]
pub fn get_cell_value_rules_nif<'a>(
    env: rustler::Env<'a>,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    range: Option<String>,
) -> Result<Vec<rustler::Term<'a>>, (Atom, String)> {
    let result = panic::catch_unwind(AssertUnwindSafe(
        || -> Result<Vec<rustler::Term<'a>>, String> {
            let spreadsheet_guard = resource.spreadsheet.lock().unwrap();
            let sheet = spreadsheet_guard
                .get_sheet_by_name(&sheet_name)
                .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

            let cf_collections = sheet.get_conditional_formatting_collection();
            let mut rules = Vec::new();

            for cf in cf_collections {
                let sqref = cf.get_sequence_of_references().get_sqref();

                // If range filter is specified, only include rules that match
                if let Some(ref filter_range) = range {
                    if sqref != *filter_range {
                        continue;
                    }
                }

                for rule in cf.get_conditional_collection() {
                    if *rule.get_type() == ConditionalFormatValues::CellIs {
                        let mut rule_map = extract_cell_value_rule(env, rule);
                        rule_map = rule_map
                            .map_put(atoms::range(), sqref.clone())
                            .ok()
                            .unwrap();
                        rules.push(rule_map);
                    }
                }
            }

            Ok(rules)
        },
    ));

    match result {
        Ok(Ok(rules)) => Ok(rules),
        Ok(Err(msg)) => Err((atoms::error(), msg)),
        Err(_) => Err((
            atoms::error(),
            "Error occurred in get_cell_value_rules".to_string(),
        )),
    }
}

#[rustler::nif(name = "get_color_scales")]
pub fn get_color_scales_nif<'a>(
    env: rustler::Env<'a>,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    range: Option<String>,
) -> Result<Vec<rustler::Term<'a>>, (Atom, String)> {
    let result = panic::catch_unwind(AssertUnwindSafe(
        || -> Result<Vec<rustler::Term<'a>>, String> {
            let spreadsheet_guard = resource.spreadsheet.lock().unwrap();
            let sheet = spreadsheet_guard
                .get_sheet_by_name(&sheet_name)
                .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

            let cf_collections = sheet.get_conditional_formatting_collection();
            let mut rules = Vec::new();

            for cf in cf_collections {
                let sqref = cf.get_sequence_of_references().get_sqref();

                // If range filter is specified, only include rules that match
                if let Some(ref filter_range) = range {
                    if sqref != *filter_range {
                        continue;
                    }
                }

                for rule in cf.get_conditional_collection() {
                    if *rule.get_type() == ConditionalFormatValues::ColorScale {
                        let mut rule_map = extract_color_scale(env, rule);
                        rule_map = rule_map
                            .map_put(atoms::range(), sqref.clone())
                            .ok()
                            .unwrap();
                        rules.push(rule_map);
                    }
                }
            }

            Ok(rules)
        },
    ));

    match result {
        Ok(Ok(rules)) => Ok(rules),
        Ok(Err(msg)) => Err((atoms::error(), msg)),
        Err(_) => Err((
            atoms::error(),
            "Error occurred in get_color_scales".to_string(),
        )),
    }
}

#[rustler::nif(name = "get_data_bars")]
pub fn get_data_bars_nif<'a>(
    env: rustler::Env<'a>,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    range: Option<String>,
) -> Result<Vec<rustler::Term<'a>>, (Atom, String)> {
    let result = panic::catch_unwind(AssertUnwindSafe(
        || -> Result<Vec<rustler::Term<'a>>, String> {
            let spreadsheet_guard = resource.spreadsheet.lock().unwrap();
            let sheet = spreadsheet_guard
                .get_sheet_by_name(&sheet_name)
                .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

            let cf_collections = sheet.get_conditional_formatting_collection();
            let mut rules = Vec::new();

            for cf in cf_collections {
                let sqref = cf.get_sequence_of_references().get_sqref();

                // If range filter is specified, only include rules that match
                if let Some(ref filter_range) = range {
                    if sqref != *filter_range {
                        continue;
                    }
                }

                for rule in cf.get_conditional_collection() {
                    if *rule.get_type() == ConditionalFormatValues::DataBar {
                        let mut rule_map = extract_data_bar(env, rule);
                        rule_map = rule_map
                            .map_put(atoms::range(), sqref.clone())
                            .ok()
                            .unwrap();
                        rules.push(rule_map);
                    }
                }
            }

            Ok(rules)
        },
    ));

    match result {
        Ok(Ok(rules)) => Ok(rules),
        Ok(Err(msg)) => Err((atoms::error(), msg)),
        Err(_) => Err((
            atoms::error(),
            "Error occurred in get_data_bars".to_string(),
        )),
    }
}

#[rustler::nif(name = "get_icon_sets")]
pub fn get_icon_sets_nif<'a>(
    env: rustler::Env<'a>,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    range: Option<String>,
) -> Result<Vec<rustler::Term<'a>>, (Atom, String)> {
    let result = panic::catch_unwind(AssertUnwindSafe(
        || -> Result<Vec<rustler::Term<'a>>, String> {
            let spreadsheet_guard = resource.spreadsheet.lock().unwrap();
            let sheet = spreadsheet_guard
                .get_sheet_by_name(&sheet_name)
                .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

            let cf_collections = sheet.get_conditional_formatting_collection();
            let mut rules = Vec::new();

            for cf in cf_collections {
                let sqref = cf.get_sequence_of_references().get_sqref();

                // If range filter is specified, only include rules that match
                if let Some(ref filter_range) = range {
                    if sqref != *filter_range {
                        continue;
                    }
                }

                for rule in cf.get_conditional_collection() {
                    if *rule.get_type() == ConditionalFormatValues::IconSet {
                        let mut rule_map = extract_icon_set(env, rule);
                        rule_map = rule_map
                            .map_put(atoms::range(), sqref.clone())
                            .ok()
                            .unwrap();
                        rules.push(rule_map);
                    }
                }
            }

            Ok(rules)
        },
    ));

    match result {
        Ok(Ok(rules)) => Ok(rules),
        Ok(Err(msg)) => Err((atoms::error(), msg)),
        Err(_) => Err((
            atoms::error(),
            "Error occurred in get_icon_sets".to_string(),
        )),
    }
}

#[rustler::nif(name = "get_top_bottom_rules")]
pub fn get_top_bottom_rules_nif<'a>(
    env: rustler::Env<'a>,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    range: Option<String>,
) -> Result<Vec<rustler::Term<'a>>, (Atom, String)> {
    let result = panic::catch_unwind(AssertUnwindSafe(
        || -> Result<Vec<rustler::Term<'a>>, String> {
            let spreadsheet_guard = resource.spreadsheet.lock().unwrap();
            let sheet = spreadsheet_guard
                .get_sheet_by_name(&sheet_name)
                .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

            let cf_collections = sheet.get_conditional_formatting_collection();
            let mut rules = Vec::new();

            for cf in cf_collections {
                let sqref = cf.get_sequence_of_references().get_sqref();

                // If range filter is specified, only include rules that match
                if let Some(ref filter_range) = range {
                    if sqref != *filter_range {
                        continue;
                    }
                }

                for rule in cf.get_conditional_collection() {
                    if *rule.get_type() == ConditionalFormatValues::Top10 {
                        let mut rule_map = extract_top_bottom_rule(env, rule);
                        rule_map = rule_map
                            .map_put(atoms::range(), sqref.clone())
                            .ok()
                            .unwrap();
                        rules.push(rule_map);
                    }
                }
            }

            Ok(rules)
        },
    ));

    match result {
        Ok(Ok(rules)) => Ok(rules),
        Ok(Err(msg)) => Err((atoms::error(), msg)),
        Err(_) => Err((
            atoms::error(),
            "Error occurred in get_top_bottom_rules".to_string(),
        )),
    }
}

#[rustler::nif(name = "get_above_below_average_rules")]
pub fn get_above_below_average_rules_nif<'a>(
    env: rustler::Env<'a>,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    range: Option<String>,
) -> Result<Vec<rustler::Term<'a>>, (Atom, String)> {
    let result = panic::catch_unwind(AssertUnwindSafe(
        || -> Result<Vec<rustler::Term<'a>>, String> {
            let spreadsheet_guard = resource.spreadsheet.lock().unwrap();
            let sheet = spreadsheet_guard
                .get_sheet_by_name(&sheet_name)
                .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

            let cf_collections = sheet.get_conditional_formatting_collection();
            let mut rules = Vec::new();

            for cf in cf_collections {
                let sqref = cf.get_sequence_of_references().get_sqref();

                // If range filter is specified, only include rules that match
                if let Some(ref filter_range) = range {
                    if sqref != *filter_range {
                        continue;
                    }
                }

                for rule in cf.get_conditional_collection() {
                    if *rule.get_type() == ConditionalFormatValues::AboveAverage {
                        let mut rule_map = extract_above_below_average_rule(env, rule);
                        rule_map = rule_map
                            .map_put(atoms::range(), sqref.clone())
                            .ok()
                            .unwrap();
                        rules.push(rule_map);
                    }
                }
            }

            Ok(rules)
        },
    ));

    match result {
        Ok(Ok(rules)) => Ok(rules),
        Ok(Err(msg)) => Err((atoms::error(), msg)),
        Err(_) => Err((
            atoms::error(),
            "Error occurred in get_above_below_average_rules".to_string(),
        )),
    }
}

#[rustler::nif(name = "get_text_rules")]
pub fn get_text_rules_nif<'a>(
    env: rustler::Env<'a>,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    range: Option<String>,
) -> Result<Vec<rustler::Term<'a>>, (Atom, String)> {
    let result = panic::catch_unwind(AssertUnwindSafe(
        || -> Result<Vec<rustler::Term<'a>>, String> {
            let spreadsheet_guard = resource.spreadsheet.lock().unwrap();
            let sheet = spreadsheet_guard
                .get_sheet_by_name(&sheet_name)
                .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

            let cf_collections = sheet.get_conditional_formatting_collection();
            let mut rules = Vec::new();

            for cf in cf_collections {
                let sqref = cf.get_sequence_of_references().get_sqref();

                // If range filter is specified, only include rules that match
                if let Some(ref filter_range) = range {
                    if sqref != *filter_range {
                        continue;
                    }
                }

                for rule in cf.get_conditional_collection() {
                    let rule_type = rule.get_type();
                    if *rule_type == ConditionalFormatValues::ContainsText
                        || *rule_type == ConditionalFormatValues::BeginsWith
                        || *rule_type == ConditionalFormatValues::EndsWith
                    {
                        let mut rule_map = extract_text_rule(env, rule);
                        rule_map = rule_map
                            .map_put(atoms::range(), sqref.clone())
                            .ok()
                            .unwrap();
                        rules.push(rule_map);
                    }
                }
            }

            Ok(rules)
        },
    ));

    match result {
        Ok(Ok(rules)) => Ok(rules),
        Ok(Err(msg)) => Err((atoms::error(), msg)),
        Err(_) => Err((
            atoms::error(),
            "Error occurred in get_text_rules".to_string(),
        )),
    }
}
