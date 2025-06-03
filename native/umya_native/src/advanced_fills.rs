use crate::{atoms, helpers, UmyaSpreadsheet};
use rustler::{Atom, Error as NifError, NifResult, ResourceArc};
use umya_spreadsheet::helper::coordinate::CellCoordinates;
use umya_spreadsheet::{EnumTrait, GradientFill, GradientStop, PatternFill, PatternValues};

// Helper function to create CellCoordinates from address string
fn cell_coordinates_from_address(address: &str) -> CellCoordinates {
    address.into()
}

// Helper function to convert pattern name from camelCase to snake_case
fn convert_pattern_name_to_snake_case(name: &str) -> String {
    let mut result = String::with_capacity(name.len() + 5); // Add some extra space for underscores

    // Special case: handle "None" separately
    if name == "None" {
        return "none".to_string();
    }

    for (i, c) in name.char_indices() {
        // If this character is uppercase and it's not the first character,
        // and the previous character is not uppercase (to handle consecutive uppercase letters correctly),
        // add an underscore before it
        if c.is_uppercase() && i > 0 && !name.chars().nth(i - 1).unwrap_or('_').is_uppercase() {
            result.push('_');
        }
        // Convert to lowercase and add to result
        result.push(c.to_lowercase().next().unwrap());
    }

    result
}

/// Set a gradient fill for a cell
#[rustler::nif]
pub fn set_gradient_fill(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    degree: f64,
    gradient_stops: Vec<(f64, String)>,
) -> NifResult<Atom> {
    // Validate that we have at least one gradient stop
    if gradient_stops.is_empty() {
        return Err(NifError::Term(Box::new((
            atoms::error(),
            "At least one gradient stop is required".to_string(),
        ))));
    }

    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Create gradient fill
            let mut gradient_fill = GradientFill::default();
            gradient_fill.set_degree(degree);

            // Add gradient stops
            for (position, color_str) in gradient_stops {
                let color_obj = match helpers::color_helper::create_color_object(&color_str) {
                    Ok(color) => color,
                    Err(_) => {
                        return Err(NifError::Term(Box::new((
                            atoms::error(),
                            format!("Invalid color format: {}", color_str),
                        ))))
                    }
                };

                let mut gradient_stop = GradientStop::default();
                gradient_stop.set_position(position);
                gradient_stop.set_color(color_obj);
                gradient_fill.get_gradient_stop_mut().push(gradient_stop);
            }

            // Convert string address to CellCoordinates
            let cell_coord = cell_coordinates_from_address(&cell_address);
            let style = sheet.get_style_mut(cell_coord);
            style.get_fill_mut().set_gradient_fill(gradient_fill);

            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Worksheet not found".to_string(),
        )))),
    }
}

/// Set a linear gradient fill for a cell (convenience function)
#[rustler::nif]
pub fn set_linear_gradient_fill(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    start_color: String,
    end_color: String,
    angle_degrees: Option<f64>,
) -> NifResult<Atom> {
    let angle_degrees = angle_degrees.unwrap_or(0.0);

    // Create gradient stops for a simple linear gradient
    let gradient_stops = vec![(0.0, start_color), (1.0, end_color)];

    // Call the main gradient fill function implementation directly
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Create gradient fill
            let mut gradient_fill = GradientFill::default();
            gradient_fill.set_degree(angle_degrees);

            // Add gradient stops
            for (position, color_str) in gradient_stops {
                let color_obj = match helpers::color_helper::create_color_object(&color_str) {
                    Ok(color) => color,
                    Err(_) => {
                        return Err(NifError::Term(Box::new((
                            atoms::error(),
                            format!("Invalid color format: {}", color_str),
                        ))))
                    }
                };

                let mut gradient_stop = GradientStop::default();
                gradient_stop.set_position(position);
                gradient_stop.set_color(color_obj);
                gradient_fill.get_gradient_stop_mut().push(gradient_stop);
            }

            // Convert string address to CellCoordinates
            let cell_coord = cell_coordinates_from_address(&cell_address);
            let style = sheet.get_style_mut(cell_coord);
            style.get_fill_mut().set_gradient_fill(gradient_fill);

            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Worksheet not found".to_string(),
        )))),
    }
}

/// Set a radial gradient fill for a cell (using a 90-degree angle for approximation)
#[rustler::nif]
pub fn set_radial_gradient_fill(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    center_color: String,
    edge_color: String,
) -> NifResult<Atom> {
    // Create gradient stops for a radial gradient (approximated as linear with 90 degrees)
    let gradient_stops = vec![(0.0, center_color), (1.0, edge_color)];

    // Call the main gradient fill function implementation directly
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Create gradient fill
            let mut gradient_fill = GradientFill::default();
            gradient_fill.set_degree(90.0);

            // Add gradient stops
            for (position, color_str) in gradient_stops {
                let color_obj = match helpers::color_helper::create_color_object(&color_str) {
                    Ok(color) => color,
                    Err(_) => {
                        return Err(NifError::Term(Box::new((
                            atoms::error(),
                            format!("Invalid color format: {}", color_str),
                        ))))
                    }
                };

                let mut gradient_stop = GradientStop::default();
                gradient_stop.set_position(position);
                gradient_stop.set_color(color_obj);
                gradient_fill.get_gradient_stop_mut().push(gradient_stop);
            }

            // Convert string address to CellCoordinates
            let cell_coord = cell_coordinates_from_address(&cell_address);
            let style = sheet.get_style_mut(cell_coord);
            style.get_fill_mut().set_gradient_fill(gradient_fill);

            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Worksheet not found".to_string(),
        )))),
    }
}

/// Set a three-color gradient fill for a cell
#[rustler::nif]
pub fn set_three_color_gradient_fill(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    start_color: String,
    middle_color: String,
    end_color: String,
    angle: Option<f64>,
) -> NifResult<Atom> {
    let angle_degrees = angle.unwrap_or(0.0);

    // Create gradient stops for a three-color gradient
    let gradient_stops = vec![(0.0, start_color), (0.5, middle_color), (1.0, end_color)];

    // Call the main gradient fill function implementation directly
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Create gradient fill
            let mut gradient_fill = GradientFill::default();
            gradient_fill.set_degree(angle_degrees);

            // Add gradient stops
            for (position, color_str) in gradient_stops {
                let color_obj = match helpers::color_helper::create_color_object(&color_str) {
                    Ok(color) => color,
                    Err(_) => {
                        return Err(NifError::Term(Box::new((
                            atoms::error(),
                            format!("Invalid color format: {}", color_str),
                        ))))
                    }
                };

                let mut gradient_stop = GradientStop::default();
                gradient_stop.set_position(position);
                gradient_stop.set_color(color_obj);
                gradient_fill.get_gradient_stop_mut().push(gradient_stop);
            }

            // Convert string address to CellCoordinates
            let cell_coord = cell_coordinates_from_address(&cell_address);
            let style = sheet.get_style_mut(cell_coord);
            style.get_fill_mut().set_gradient_fill(gradient_fill);

            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Worksheet not found".to_string(),
        )))),
    }
}

/// Set a custom multi-stop gradient fill for a cell
#[rustler::nif]
pub fn set_custom_gradient_fill(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    angle: f64,
    stops: Vec<(f64, String)>,
    validate_positions: Option<bool>,
) -> NifResult<Atom> {
    // Validate gradient stops
    if stops.is_empty() {
        return Err(NifError::Term(Box::new((
            atoms::error(),
            "At least one gradient stop is required".to_string(),
        ))));
    }

    // Validate position values if validation is enabled (default: true)
    if validate_positions.unwrap_or(true) {
        for (position, _) in &stops {
            if *position < 0.0 || *position > 1.0 {
                return Err(NifError::Term(Box::new((
                    atoms::error(),
                    format!(
                        "Gradient stop position must be between 0.0 and 1.0, got: {}",
                        position
                    ),
                ))));
            }
        }
    }

    // Call the main gradient fill function implementation directly
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Create gradient fill
            let mut gradient_fill = GradientFill::default();
            gradient_fill.set_degree(angle);

            // Add gradient stops
            for (position, color_str) in stops {
                let color_obj = match helpers::color_helper::create_color_object(&color_str) {
                    Ok(color) => color,
                    Err(_) => {
                        return Err(NifError::Term(Box::new((
                            atoms::error(),
                            format!("Invalid color format: {}", color_str),
                        ))))
                    }
                };

                let mut gradient_stop = GradientStop::default();
                gradient_stop.set_position(position);
                gradient_stop.set_color(color_obj);
                gradient_fill.get_gradient_stop_mut().push(gradient_stop);
            }

            // Convert string address to CellCoordinates
            let cell_coord = cell_coordinates_from_address(&cell_address);
            let style = sheet.get_style_mut(cell_coord);
            style.get_fill_mut().set_gradient_fill(gradient_fill);

            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Worksheet not found".to_string(),
        )))),
    }
}

/// Get gradient fill information from a cell
#[rustler::nif]
pub fn get_gradient_fill(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<(f64, Vec<(f64, String)>)> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            // Convert string address to CellCoordinates
            let cell_coord = cell_coordinates_from_address(&cell_address);
            if let Some(cell) = sheet.get_cell(cell_coord) {
                if let Some(fill) = cell.get_style().get_fill() {
                    if let Some(gradient_fill) = fill.get_gradient_fill() {
                        let degree = *gradient_fill.get_degree();
                        let stops: Vec<(f64, String)> = gradient_fill
                            .get_gradient_stop()
                            .iter()
                            .map(|stop| {
                                (
                                    *stop.get_position(),
                                    stop.get_color().get_argb().to_string(),
                                )
                            })
                            .collect();

                        return Ok((degree, stops));
                    }
                }
            }
            // Return error if no gradient fill found
            Err(NifError::Term(Box::new((
                atoms::error(),
                "No gradient fill found for cell".to_string(),
            ))))
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Worksheet not found".to_string(),
        )))),
    }
}

/// Set a pattern fill for a cell
#[rustler::nif]
pub fn set_pattern_fill(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    pattern_type: String,
    foreground_color: Option<String>,
    background_color: Option<String>,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Convert string address to CellCoordinates
            let cell_coord = cell_coordinates_from_address(&cell_address);
            let style = sheet.get_style_mut(cell_coord);

            // Set pattern type from string
            // Convert pattern_type to a normalized format (remove underscores, hyphens, and lowercase)
            let normalized_pattern_type = pattern_type
                .to_lowercase()
                .replace("_", "")
                .replace("-", "");

            let pattern_value = match normalized_pattern_type.as_str() {
                "solid" => PatternValues::Solid,
                "none" => PatternValues::None,
                "gray125" => PatternValues::Gray125,
                "darkgray" => PatternValues::DarkGray,
                "mediumgray" => PatternValues::MediumGray,
                "lightgray" => PatternValues::LightGray,
                "gray0625" => PatternValues::Gray0625,
                "darkdown" => PatternValues::DarkDown,
                "darkgrid" => PatternValues::DarkGrid,
                "darkhorizontal" => PatternValues::DarkHorizontal,
                "darktrellis" => PatternValues::DarkTrellis,
                "darkup" => PatternValues::DarkUp,
                "darkvertical" => PatternValues::DarkVertical,
                "lightdown" => PatternValues::LightDown,
                "lightgrid" => PatternValues::LightGrid,
                "lighthorizontal" => PatternValues::LightHorizontal,
                "lighttrellis" => PatternValues::LightTrellis,
                "lightup" => PatternValues::LightUp,
                "lightvertical" => PatternValues::LightVertical,
                _ => {
                    return Err(NifError::Term(Box::new((
                        atoms::error(),
                        format!("Invalid pattern type: {}", pattern_type),
                    ))))
                }
            };

            let mut pattern_fill = PatternFill::default();
            pattern_fill.set_pattern_type(pattern_value);
            style.get_fill_mut().set_pattern_fill(pattern_fill);

            // Set foreground color if provided
            if let Some(fg_color_str) = foreground_color {
                let fg_color = match helpers::color_helper::create_color_object(&fg_color_str) {
                    Ok(color) => color,
                    Err(_) => {
                        return Err(NifError::Term(Box::new((
                            atoms::error(),
                            format!("Invalid foreground color format: {}", fg_color_str),
                        ))))
                    }
                };
                let fill = style.get_fill_mut();
                fill.get_pattern_fill_mut().set_foreground_color(fg_color);
            }

            // Set background color if provided
            if let Some(bg_color_str) = background_color {
                let bg_color = match helpers::color_helper::create_color_object(&bg_color_str) {
                    Ok(color) => color,
                    Err(_) => {
                        return Err(NifError::Term(Box::new((
                            atoms::error(),
                            format!("Invalid background color format: {}", bg_color_str),
                        ))))
                    }
                };
                let fill = style.get_fill_mut();
                fill.get_pattern_fill_mut().set_background_color(bg_color);
            }

            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Worksheet not found".to_string(),
        )))),
    }
}

/// Get pattern fill information from a cell
#[rustler::nif]
pub fn get_pattern_fill(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<(String, Option<String>, Option<String>)> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            // Convert string address to CellCoordinates
            let cell_coord = cell_coordinates_from_address(&cell_address);
            if let Some(cell) = sheet.get_cell(cell_coord) {
                if let Some(fill) = cell.get_style().get_fill() {
                    if let Some(pattern_fill) = fill.get_pattern_fill() {
                        // Get the pattern type as a string
                        let raw_pattern_name = pattern_fill.get_pattern_type().get_value_string();
                        // Convert camelCase to snake_case for consistency with Elixir
                        let pattern_name = convert_pattern_name_to_snake_case(&raw_pattern_name);

                        // If the pattern is "none", return an error since it's not a real pattern fill
                        if pattern_name == "none" {
                            return Err(NifError::Term(Box::new((
                                atoms::error(),
                                "No pattern fill found for cell".to_string(),
                            ))));
                        }

                        let fg_color = pattern_fill
                            .get_foreground_color()
                            .map(|color| color.get_argb().to_string());

                        let bg_color = pattern_fill
                            .get_background_color()
                            .map(|color| color.get_argb().to_string());

                        return Ok((pattern_name, fg_color, bg_color));
                    }
                }
            }
            // Return error if no pattern fill found
            Err(NifError::Term(Box::new((
                atoms::error(),
                "No pattern fill found for cell".to_string(),
            ))))
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Worksheet not found".to_string(),
        )))),
    }
}

/// Clear all fills from a cell
#[rustler::nif]
pub fn clear_fill(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Convert string address to CellCoordinates
            let cell_coord = cell_coordinates_from_address(&cell_address);

            // Get the style and reset the fill to a basic pattern fill (None)
            let style = sheet.get_style_mut(cell_coord);
            let fill = style.get_fill_mut();

            // Create a new empty pattern fill
            let mut pattern_fill = PatternFill::default();
            pattern_fill.set_pattern_type(PatternValues::None);
            fill.set_pattern_fill(pattern_fill);

            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Worksheet not found".to_string(),
        )))),
    }
}
