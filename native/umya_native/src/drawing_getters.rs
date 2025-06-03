use rustler::{Encoder, Env, ResourceArc, Term};
use std::panic::{self, AssertUnwindSafe};
use umya_spreadsheet::helper::coordinate::string_from_column_index;
use umya_spreadsheet::structs::drawing::spreadsheet::Shape;
use umya_spreadsheet::Spreadsheet as UmyaSpreadsheetInternal;

use crate::atoms::{self, error, not_found, ok};
use crate::helpers::cell_helpers::{check_inside, is_cell_in_range};
use crate::UmyaSpreadsheet;

// Function to get all shapes from a worksheet
#[rustler::nif]
pub fn get_shapes_nif(
    env: Env,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_range: Option<String>,
) -> Term {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let spreadsheet = &resource.spreadsheet;
        let guard = spreadsheet.lock().map_err(|_| error())?;
        let spreadsheet_ref: &UmyaSpreadsheetInternal = &guard;

        // Try to get the worksheet
        let worksheet = match spreadsheet_ref.get_sheet_by_name(&sheet_name) {
            Some(ws) => ws,
            None => return Err(not_found()),
        };

        // Get the drawing collection
        let mut result_shapes = Vec::new();

        // Process shapes from one cell anchors
        let drawing = worksheet.get_worksheet_drawing();
        for anchor in drawing.get_one_cell_anchor_collection().iter() {
            if let Some(shape) = anchor.get_shape() {
                // Skip text boxes (they'll be handled separately)
                if shape.get_text_body().is_none() {
                    if let Some(cell_range_str) = &cell_range {
                        // Filter by cell range if provided
                        let cell_marker = anchor.get_from_marker();
                        let col = *cell_marker.get_col();
                        let row = *cell_marker.get_row();
                        let cell_address =
                            format!("{}{}", string_from_column_index(&(col + 1)), row + 1);

                        // Check if cell is within the specified range
                        if !is_cell_in_range(&cell_address, cell_range_str) {
                            continue;
                        }
                    }

                    // Determine shape type
                    let shape_type = get_shape_type_str(shape);

                    // Get shape properties
                    let from_marker = anchor.get_from_marker();
                    let from_cell = format!(
                        "{}{}",
                        string_from_column_index(&(from_marker.get_col() + 1)),
                        from_marker.get_row() + 1
                    );

                    // Get dimensions from shape's transform instead of anchor extent
                    let shape_props = shape.get_shape_properties();
                    let (width, height) = if let Some(transform) = shape_props.get_transform2d() {
                        let extents = transform.get_extents();
                        (*extents.get_cx() as f64, *extents.get_cy() as f64)
                    } else {
                        // Fallback to default values if no transform
                        (100.0, 50.0)
                    };

                    // Get fill color and outline properties
                    let (fill_color, outline_color, outline_width) =
                        get_shape_style_properties(shape);

                    let mut shape_map = rustler::types::map::map_new(env);
                    shape_map = shape_map.map_put(atoms::type_(), shape_type.encode(env)).ok().unwrap();
                    shape_map = shape_map.map_put(atoms::cell(), from_cell.encode(env)).ok().unwrap();
                    shape_map = shape_map.map_put(atoms::width(), width.encode(env)).ok().unwrap();
                    shape_map = shape_map.map_put(atoms::height(), height.encode(env)).ok().unwrap();
                    shape_map = shape_map.map_put(atoms::fill_color(), fill_color.encode(env)).ok().unwrap();
                    shape_map = shape_map.map_put(atoms::outline_color(), outline_color.encode(env)).ok().unwrap();
                    shape_map = shape_map.map_put(atoms::outline_width(), outline_width.encode(env)).ok().unwrap();

                    result_shapes.push(shape_map);
                }
            }
        }

        Ok((ok(), result_shapes))
    }));

    match result {
        Ok(Ok((status, shapes))) => (status, shapes).encode(env),
        Ok(Err(err)) => (error(), err).encode(env),
        Err(_) => (error(), error()).encode(env),
    }
}

// Function to get all text boxes from a worksheet
#[rustler::nif]
pub fn get_text_boxes_nif(
    env: Env,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_range: Option<String>,
) -> Term {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let spreadsheet = &resource.spreadsheet;
        let guard = spreadsheet.lock().map_err(|_| error())?;
        let spreadsheet_ref: &UmyaSpreadsheetInternal = &guard;

        // Try to get the worksheet
        let worksheet = match spreadsheet_ref.get_sheet_by_name(&sheet_name) {
            Some(ws) => ws,
            None => return Err(not_found()),
        };

        // Get the drawing collection
        let mut result_text_boxes = Vec::new();

        // Process text boxes from one cell anchors
        let drawing = worksheet.get_worksheet_drawing();
        for anchor in drawing.get_one_cell_anchor_collection().iter() {
            if let Some(shape) = anchor.get_shape() {
                // Only look for shapes with text body (text boxes)
                if let Some(text_body) = shape.get_text_body() {
                    if let Some(cell_range_str) = &cell_range {
                        // Filter by cell range if provided
                        let cell_marker = anchor.get_from_marker();
                        let col = *cell_marker.get_col();
                        let row = *cell_marker.get_row();
                        let cell_address =
                            format!("{}{}", string_from_column_index(&(col + 1)), row + 1);

                        // Check if cell is within the specified range
                        if !is_cell_in_range(&cell_address, cell_range_str) {
                            continue;
                        }
                    }

                    // Get text box properties
                    let from_marker = anchor.get_from_marker();
                    let from_cell = format!(
                        "{}{}",
                        string_from_column_index(&(from_marker.get_col() + 1)),
                        from_marker.get_row() + 1
                    );

                    // Get dimensions from shape's transform instead of anchor extent
                    let shape_props = shape.get_shape_properties();
                    let (width, height) = if let Some(transform) = shape_props.get_transform2d() {
                        let extents = transform.get_extents();
                        (*extents.get_cx() as f64, *extents.get_cy() as f64)
                    } else {
                        // Fallback to default values if no transform
                        (150.0, 75.0)
                    };

                    // Get text content - use the array of paragraphs directly
                    let paragraphs = text_body.get_paragraph();
                    let text = if !paragraphs.is_empty() {
                        // Get text from the first paragraph
                        paragraphs[0]
                            .get_run()
                            .iter()
                            .map(|tr| tr.get_text().to_string())
                            .collect::<Vec<String>>()
                            .join("")
                    } else {
                        String::new()
                    };

                    // Get fill color, outline properties, and text color
                    let (fill_color, outline_color, outline_width) =
                        get_shape_style_properties(shape);

                    // Get the text color from the first text run of the first paragraph if available
                    let text_color = if !paragraphs.is_empty() {
                        if let Some(text_run) = paragraphs[0].get_run().first() {
                            let run_properties = text_run.get_run_properties();
                            if let Some(solid_fill) = run_properties.get_solid_fill() {
                                if let Some(rgb_color) = solid_fill.get_rgb_color_model_hex() {
                                    rgb_color.get_val().to_string()
                                } else {
                                    "000000".to_string() // Default to black
                                }
                            } else {
                                "000000".to_string()
                            }
                        } else {
                            "000000".to_string()
                        }
                    } else {
                        "000000".to_string()
                    };

                    let mut text_box_map = rustler::types::map::map_new(env);
                    text_box_map = text_box_map.map_put(atoms::cell(), from_cell.encode(env)).ok().unwrap();
                    text_box_map = text_box_map.map_put(atoms::text(), text.encode(env)).ok().unwrap();
                    text_box_map = text_box_map.map_put(atoms::width(), width.encode(env)).ok().unwrap();
                    text_box_map = text_box_map.map_put(atoms::height(), height.encode(env)).ok().unwrap();
                    text_box_map = text_box_map.map_put(atoms::fill_color(), fill_color.encode(env)).ok().unwrap();
                    text_box_map = text_box_map.map_put(atoms::text_color(), text_color.encode(env)).ok().unwrap();
                    text_box_map = text_box_map.map_put(atoms::outline_color(), outline_color.encode(env)).ok().unwrap();
                    text_box_map = text_box_map.map_put(atoms::outline_width(), outline_width.encode(env)).ok().unwrap();

                    result_text_boxes.push(text_box_map);
                }
            }
        }

        Ok((ok(), result_text_boxes))
    }));

    match result {
        Ok(Ok((status, text_boxes))) => (status, text_boxes).encode(env),
        Ok(Err(err)) => (error(), err).encode(env),
        Err(_) => (error(), error()).encode(env),
    }
}

// Function to get all connectors from a worksheet
#[rustler::nif]
pub fn get_connectors_nif(
    env: Env,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_range: Option<String>,
) -> Term {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let spreadsheet = &resource.spreadsheet;
        let mut guard = spreadsheet.lock().map_err(|_| error())?;
        let spreadsheet_ref: &mut UmyaSpreadsheetInternal = &mut guard;

        // Try to get the worksheet
        let worksheet = match spreadsheet_ref.get_sheet_by_name(&sheet_name) {
            Some(ws) => ws,
            None => return Err(not_found()),
        };

        // Get the drawing collection
        let mut result_connectors = Vec::new();

        // Process connectors from two cell anchors
        let drawing = worksheet.get_worksheet_drawing();
        for anchor in drawing.get_two_cell_anchor_collection().iter() {
            if let Some(shape) = anchor.get_shape() {
                // Filter for connector types
                if is_connector_shape(shape) {
                    let from_marker = anchor.get_from_marker();
                    let to_marker = anchor.get_to_marker(); // Adjusted method name

                    let from_cell = format!(
                        "{}{}",
                        string_from_column_index(&(from_marker.get_col() + 1)),
                        from_marker.get_row() + 1
                    );

                    let to_cell = format!(
                        "{}{}",
                        string_from_column_index(&(to_marker.get_col() + 1)),
                        to_marker.get_row() + 1
                    );

                    // If cell range filter is provided
                    if let Some(cell_range_str) = &cell_range {
                        // Check if either end of the connector is in the range
                        if !check_inside(&from_cell, cell_range_str)
                            && !check_inside(&to_cell, cell_range_str)
                        {
                            continue;
                        }
                    }

                    // Get line color and width
                    let (_, line_color, line_width) = get_shape_style_properties(shape);

                    let mut connector_map = rustler::types::map::map_new(env);
                    connector_map = connector_map.map_put(atoms::from_cell(), from_cell.encode(env)).ok().unwrap();
                    connector_map = connector_map.map_put(atoms::to_cell(), to_cell.encode(env)).ok().unwrap();
                    connector_map = connector_map.map_put(atoms::line_color(), line_color.encode(env)).ok().unwrap();
                    connector_map = connector_map.map_put(atoms::line_width(), line_width.encode(env)).ok().unwrap();

                    result_connectors.push(connector_map);
                }
            }
        }

        Ok((ok(), result_connectors))
    }));

    match result {
        Ok(Ok((status, connectors))) => (status, connectors).encode(env),
        Ok(Err(err)) => (error(), err).encode(env),
        Err(_) => (error(), error()).encode(env),
    }
}

// Function to check if a sheet has any drawing objects
#[rustler::nif]
pub fn has_drawing_objects_nif(
    env: Env,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_range: Option<String>,
) -> Term {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let spreadsheet = &resource.spreadsheet;
        let mut guard = spreadsheet.lock().map_err(|_| error())?;
        let spreadsheet_ref: &mut UmyaSpreadsheetInternal = &mut guard;

        // Try to get the worksheet
        let worksheet = match spreadsheet_ref.get_sheet_by_name(&sheet_name) {
            Some(ws) => ws,
            None => return Err(not_found()),
        };

        // Check if the worksheet has drawing collection
        let drawing = worksheet.get_worksheet_drawing();
        // If no cell range is specified, just check if there are any anchors
        if cell_range.is_none() {
            if !drawing.get_one_cell_anchor_collection().is_empty()
                || !drawing.get_two_cell_anchor_collection().is_empty()
            {
                return Ok((ok(), true));
            }
            return Ok((ok(), false));
        }

        // If cell range is provided, check if any drawing is in that range
        let cell_range_str = cell_range.unwrap();

        // Check one cell anchors
        for anchor in drawing.get_one_cell_anchor_collection().iter() {
            let cell_marker = anchor.get_from_marker();
            let col = *cell_marker.get_col();
            let row = *cell_marker.get_row();
            let cell_address = format!("{}{}", string_from_column_index(&(col + 1)), row + 1);

            if check_inside(&cell_address, &cell_range_str) {
                return Ok((ok(), true));
            }
        }

        // Check two cell anchors
        for anchor in drawing.get_two_cell_anchor_collection().iter() {
            let from_marker = anchor.get_from_marker();
            let to_marker = anchor.get_to_marker();

            let from_cell = format!(
                "{}{}",
                string_from_column_index(&(from_marker.get_col() + 1)),
                from_marker.get_row() + 1
            );

            let to_cell = format!(
                "{}{}",
                string_from_column_index(&(to_marker.get_col() + 1)),
                to_marker.get_row() + 1
            );

            if check_inside(&from_cell, &cell_range_str) || check_inside(&to_cell, &cell_range_str)
            {
                return Ok((ok(), true));
            }
        }

        Ok((ok(), false))
    }));

    match result {
        Ok(Ok((status, has_drawings))) => (status, has_drawings).encode(env),
        Ok(Err(err)) => (error(), err).encode(env),
        Err(_) => (error(), error()).encode(env),
    }
}

// Function to count drawing objects in a sheet
#[rustler::nif]
pub fn count_drawing_objects_nif(
    env: Env,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_range: Option<String>,
) -> Term {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let spreadsheet = &resource.spreadsheet;
        let mut guard = spreadsheet.lock().map_err(|_| error())?;
        let spreadsheet_ref: &mut UmyaSpreadsheetInternal = &mut guard;

        // Try to get the worksheet
        let worksheet = match spreadsheet_ref.get_sheet_by_name(&sheet_name) {
            Some(ws) => ws,
            None => return Err(not_found()),
        };

        let mut count = 0;

        // Check if the worksheet has drawing collection
        let drawing = worksheet.get_worksheet_drawing();
        // If no cell range is specified, just count all anchors
        if cell_range.is_none() {
            count += drawing.get_one_cell_anchor_collection().len();
            count += drawing.get_two_cell_anchor_collection().len();
            return Ok((ok(), count));
        }

        // If cell range is provided, count drawings in that range
        let cell_range_str = cell_range.unwrap();

        // Count one cell anchors in range
        for anchor in drawing.get_one_cell_anchor_collection().iter() {
            let cell_marker = anchor.get_from_marker();
            let col = *cell_marker.get_col();
            let row = *cell_marker.get_row();
            let cell_address = format!("{}{}", string_from_column_index(&(col + 1)), row + 1);

            if check_inside(&cell_address, &cell_range_str) {
                count += 1;
            }
        }

        // Count two cell anchors in range
        for anchor in drawing.get_two_cell_anchor_collection().iter() {
            let from_marker = anchor.get_from_marker();
            let to_marker = anchor.get_to_marker();

            let from_cell = format!(
                "{}{}",
                string_from_column_index(&(from_marker.get_col() + 1)),
                from_marker.get_row() + 1
            );

            let to_cell = format!(
                "{}{}",
                string_from_column_index(&(to_marker.get_col() + 1)),
                to_marker.get_row() + 1
            );

            if check_inside(&from_cell, &cell_range_str) || check_inside(&to_cell, &cell_range_str)
            {
                count += 1;
            }
        }

        Ok((ok(), count))
    }));

    match result {
        Ok(Ok((status, count))) => (status, count).encode(env),
        Ok(Err(err)) => (error(), err).encode(env),
        Err(_) => (error(), error()).encode(env),
    }
}

// Helper function to determine the shape type from shape properties
fn get_shape_type_str(shape: &Shape) -> String {
    let shape_properties = shape.get_shape_properties();
    let geometry = shape_properties.get_geometry();

    // Get the geometry string
    let geometry_str = geometry.get_geometry();

    match geometry_str {
        "rect" => {
            if shape.get_text_body().is_some() {
                "text_box".to_string()
            } else {
                "rectangle".to_string()
            }
        }
        "roundRect" => "rounded_rectangle".to_string(),
        "ellipse" => "ellipse".to_string(),
        "triangle" => "triangle".to_string(),
        "rtTriangle" => "right_triangle".to_string(),
        "pentagon" => "pentagon".to_string(),
        "hexagon" => "hexagon".to_string(),
        "octagon" => "octagon".to_string(),
        "trapezoid" => "trapezoid".to_string(),
        "diamond" => "diamond".to_string(),
        "rightArrow" => "arrow".to_string(),
        "line" => "line".to_string(),
        "straightConnector1" => "connector".to_string(),
        other => other.to_string(),
    }
}

// Helper function to check if a shape is a connector
fn is_connector_shape(shape: &Shape) -> bool {
    let shape_properties = shape.get_shape_properties();
    let geometry = shape_properties.get_geometry();

    // Get the geometry string and check if it's a connector
    geometry.get_geometry() == "straightConnector1"
}

// Helper function to extract style properties from a shape
fn get_shape_style_properties(shape: &Shape) -> (String, String, f64) {
    let mut fill_color = "FFFFFF".to_string(); // Default to white
    let mut outline_color = "000000".to_string(); // Default to black
    let mut outline_width = 1.0; // Default width

    let shape_properties = shape.get_shape_properties();
    // Get fill color
    if let Some(solid_fill) = shape_properties.get_solid_fill() {
        if let Some(rgb_color) = solid_fill.get_rgb_color_model_hex() {
            fill_color = rgb_color.get_val().to_string();
        }
    }

    // Get outline properties
    if let Some(outline) = shape_properties.get_outline() {
        // Get outline color
        if let Some(solid_fill) = outline.get_solid_fill() {
            if let Some(rgb_color) = solid_fill.get_rgb_color_model_hex() {
                outline_color = rgb_color.get_val().to_string();
            }
        }

        // Get outline width - directly use the value, no longer an Option
        let width = outline.get_width();
        outline_width = *width as f64 / 12700.0; // Convert EMU back to points
    }

    (fill_color, outline_color, outline_width)
}
