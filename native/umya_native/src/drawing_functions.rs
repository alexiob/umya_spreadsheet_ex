use rustler::{Atom, Error as NifError, NifResult, ResourceArc};
use std::panic::{self, AssertUnwindSafe};
use umya_spreadsheet::drawing::Transform2D;
use umya_spreadsheet::structs::drawing::spreadsheet::{
    MarkerType as CellMarker, OneCellAnchor, Shape, ShapeProperties, TwoCellAnchor,
};
use umya_spreadsheet::structs::drawing::{
    AdjustValueList, Outline, Point2DType as Point2D, PositiveFixedPercentageType,
    PositiveSize2DType, PresetGeometry, RgbColorModelHex, SolidFill,
};

use crate::atoms::{error, ok};
use crate::UmyaSpreadsheet;

// The shape types we support in the Elixir wrapper
#[derive(Debug)]
enum ShapeType {
    Rectangle,
    Ellipse,
    RoundedRectangle,
    TextBox,
    Line,
    Triangle,
    RightTriangle,
    Pentagon,
    Hexagon,
    Octagon,
    Trapezoid,
    Diamond,
    Arrow,
    Connector,
}

impl ShapeType {
    fn from_string(shape_type: &str) -> Result<ShapeType, &'static str> {
        match shape_type.to_lowercase().as_str() {
            "rectangle" => Ok(ShapeType::Rectangle),
            "ellipse" | "oval" | "circle" => Ok(ShapeType::Ellipse),
            "rounded_rectangle" => Ok(ShapeType::RoundedRectangle),
            "text_box" => Ok(ShapeType::TextBox),
            "line" => Ok(ShapeType::Line),
            "triangle" => Ok(ShapeType::Triangle),
            "right_triangle" => Ok(ShapeType::RightTriangle),
            "pentagon" => Ok(ShapeType::Pentagon),
            "hexagon" => Ok(ShapeType::Hexagon),
            "octagon" => Ok(ShapeType::Octagon),
            "trapezoid" => Ok(ShapeType::Trapezoid),
            "diamond" => Ok(ShapeType::Diamond),
            "arrow" => Ok(ShapeType::Arrow),
            "connector" => Ok(ShapeType::Connector),
            _ => Err("Unsupported shape type"),
        }
    }

    fn to_preset_geometry(&self) -> PresetGeometry {
        let mut geometry = PresetGeometry::default();

        // Set the preset geometry shape type
        match self {
            ShapeType::Rectangle => geometry.set_geometry("rect"),
            ShapeType::Ellipse => geometry.set_geometry("ellipse"),
            ShapeType::RoundedRectangle => geometry.set_geometry("roundRect"),
            ShapeType::TextBox => geometry.set_geometry("rect"), // Text box uses rectangle shape
            ShapeType::Line => geometry.set_geometry("line"),
            ShapeType::Triangle => geometry.set_geometry("triangle"),
            ShapeType::RightTriangle => geometry.set_geometry("rtTriangle"),
            ShapeType::Pentagon => geometry.set_geometry("pentagon"),
            ShapeType::Hexagon => geometry.set_geometry("hexagon"),
            ShapeType::Octagon => geometry.set_geometry("octagon"),
            ShapeType::Trapezoid => geometry.set_geometry("trapezoid"),
            ShapeType::Diamond => geometry.set_geometry("diamond"),
            ShapeType::Arrow => geometry.set_geometry("rightArrow"),
            ShapeType::Connector => geometry.set_geometry("straightConnector1"),
        };

        // For shapes that need adjust values
        match self {
            ShapeType::RoundedRectangle => {
                let adjust_list = AdjustValueList::default();
                // For now, we don't use the shape guide for rounded rectangle
                // in this simplified implementation
                geometry.set_adjust_value_list(adjust_list);
            }
            _ => {}
        }

        geometry
    }
}

// Add a new shape to a worksheet
#[rustler::nif]
pub fn add_shape(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    shape_type: String,
    width: f64,
    height: f64,
    fill_color: String,
    outline_color: String,
    outline_width: f64,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet_guard = match resource.spreadsheet.lock() {
            Ok(guard) => guard,
            Err(_) => return Err("Failed to lock spreadsheet"),
        };

        let worksheet = match spreadsheet_guard.get_sheet_by_name_mut(&sheet_name) {
            Some(ws) => ws,
            None => return Err("Sheet not found"),
        };

        // Parse the shape type
        let shape_type = match ShapeType::from_string(&shape_type) {
            Ok(shape_type) => shape_type,
            Err(msg) => return Err(msg),
        };

        // Create the shape
        let shape = create_shape(
            &cell_address,
            &shape_type,
            width,
            height,
            &fill_color,
            &outline_color,
            outline_width,
        );

        // Add the shape to the worksheet
        worksheet
            .get_worksheet_drawing_mut()
            .add_one_cell_anchor_collection(shape);

        Ok(ok())
    }));

    match result {
        Ok(Ok(atom)) => Ok(atom),
        Ok(Err(msg)) => Err(NifError::Term(Box::new((error(), msg.to_string())))),
        Err(_) => Err(NifError::Term(Box::new((
            error(),
            "Error occurred in add_shape".to_string(),
        )))),
    }
}

// Add a text box to a worksheet
#[rustler::nif]
pub fn add_text_box(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    text: String,
    width: f64,
    height: f64,
    fill_color: String,
    text_color: String,
    outline_color: String,
    outline_width: f64,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet_guard = match resource.spreadsheet.lock() {
            Ok(guard) => guard,
            Err(_) => return Err("Failed to lock spreadsheet"),
        };

        let worksheet = match spreadsheet_guard.get_sheet_by_name_mut(&sheet_name) {
            Some(ws) => ws,
            None => return Err("Sheet not found"),
        };

        // Create a text box (which is a shape with text)
        let mut shape = create_shape(
            &cell_address,
            &ShapeType::TextBox,
            width,
            height,
            &fill_color,
            &outline_color,
            outline_width,
        );

        // Add text to the shape if it has a shape object
        if let Some(shape_ref) = shape.get_shape_mut() {
            // Create text body
            let mut text_body =
                umya_spreadsheet::structs::drawing::spreadsheet::TextBody::default();

            // Add paragraph with text
            let mut paragraph = umya_spreadsheet::structs::drawing::Paragraph::default();
            let mut run = umya_spreadsheet::structs::drawing::Run::default();

            // Set text
            run.set_text(text);

            // Set text color
            let rgb_text_color = parse_color(&text_color);
            let mut run_properties = umya_spreadsheet::structs::drawing::RunProperties::default();
            let mut text_color_fill = SolidFill::default();
            text_color_fill.set_rgb_color_model_hex(rgb_text_color);
            run_properties.set_solid_fill(text_color_fill);
            run.set_run_properties(run_properties);

            // Build the text structure
            paragraph.add_run(run);
            text_body.add_paragraph(paragraph);

            // Set the text body to the shape
            shape_ref.set_text_body(text_body);
        }

        // Add the text box to the worksheet
        worksheet
            .get_worksheet_drawing_mut()
            .add_one_cell_anchor_collection(shape);

        Ok(ok())
    }));

    match result {
        Ok(Ok(atom)) => Ok(atom),
        Ok(Err(msg)) => Err(NifError::Term(Box::new((error(), msg.to_string())))),
        Err(_) => Err(NifError::Term(Box::new((
            error(),
            "Error occurred in add_text_box".to_string(),
        )))),
    }
}

// Add a connector line between two cells
#[rustler::nif]
pub fn add_connector(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    from_cell: String,
    to_cell: String,
    line_color: String,
    line_width: f64,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet_guard = match resource.spreadsheet.lock() {
            Ok(guard) => guard,
            Err(_) => return Err("Failed to lock spreadsheet"),
        };

        let worksheet = match spreadsheet_guard.get_sheet_by_name_mut(&sheet_name) {
            Some(ws) => ws,
            None => return Err("Sheet not found"),
        };

        // Create a connector (special line that connects two points)
        let connector = create_connector(&from_cell, &to_cell, &line_color, line_width);

        // Add the connector to the worksheet
        worksheet
            .get_worksheet_drawing_mut()
            .add_two_cell_anchor_collection(connector);

        Ok(ok())
    }));

    match result {
        Ok(Ok(atom)) => Ok(atom),
        Ok(Err(msg)) => Err(NifError::Term(Box::new((error(), msg.to_string())))),
        Err(_) => Err(NifError::Term(Box::new((
            error(),
            "Error occurred in add_connector".to_string(),
        )))),
    }
}

// Helper function to create a shape
fn create_shape(
    cell_address: &str,
    shape_type: &ShapeType,
    width: f64,
    height: f64,
    fill_color: &str,
    outline_color: &str,
    outline_width: f64,
) -> OneCellAnchor {
    // Create the cell marker (position)
    let from = get_cell_marker(cell_address);

    // Create the shape with proper dimensions and styles
    let mut shape_props = ShapeProperties::default();

    // Set the geometry type
    let preset_geometry = shape_type.to_preset_geometry();
    shape_props.set_geometry(preset_geometry);

    // Set fill color
    let rgb_fill = parse_color(fill_color);
    let mut solid_fill = SolidFill::default();
    solid_fill.set_rgb_color_model_hex(rgb_fill);
    shape_props.set_solid_fill(solid_fill);

    // Set outline
    let rgb_outline = parse_color(outline_color);
    let mut outline = Outline::default();

    // Create a SolidFill for the outline color
    let mut outline_fill = SolidFill::default();
    outline_fill.set_rgb_color_model_hex(rgb_outline);
    outline.set_solid_fill(outline_fill);

    outline.set_width((outline_width * 12700.0) as u32); // Convert points to EMU for width
    shape_props.set_outline(outline);

    // Create transformation matrix
    let mut transform = Transform2D::default();

    // Create Point2D for offset manually
    let mut offset = Point2D::default();
    offset.set_x(0); // Use integer values
    offset.set_y(0);

    // Create PositiveSize2DType for extents manually
    let mut extents = PositiveSize2DType::default();
    extents.set_cx(width as i64); // Convert to i64
    extents.set_cy(height as i64);

    transform.set_offset(offset);
    transform.set_extents(extents);
    shape_props.set_transform2d(transform);

    // Create the shape
    let mut shape = Shape::default();
    shape.set_shape_properties(shape_props);

    // Create the final anchored shape
    let mut anchor = OneCellAnchor::default();
    anchor.set_from_marker(from);
    anchor.set_shape(shape);

    anchor
}

// Helper function to create a connector
fn create_connector(
    from_cell: &str,
    to_cell: &str,
    line_color: &str,
    line_width: f64,
) -> TwoCellAnchor {
    // Create marker for start and end cells
    let from = get_cell_marker(from_cell);
    let to = get_cell_marker(to_cell);

    // Create a connector shape
    let mut shape_props = ShapeProperties::default();

    // Set the connector geometry
    let mut preset_geometry = PresetGeometry::default();
    preset_geometry.set_geometry("straightConnector1");
    shape_props.set_geometry(preset_geometry);

    // Set line color and width
    let rgb_line = parse_color(line_color);
    let mut outline = Outline::default();

    // Create a SolidFill for the outline color
    let mut outline_fill = SolidFill::default();
    outline_fill.set_rgb_color_model_hex(rgb_line);
    outline.set_solid_fill(outline_fill);

    outline.set_width((line_width * 12700.0) as u32); // Convert points to EMU
    shape_props.set_outline(outline);

    // No fill for connector
    let mut solid_fill = SolidFill::default();
    solid_fill.set_rgb_color_model_hex(parse_color("transparent"));
    shape_props.set_solid_fill(solid_fill);

    // Create the connector shape
    let mut shape = Shape::default();
    shape.set_shape_properties(shape_props);

    // Create the two-cell anchor
    let mut anchor = TwoCellAnchor::default();
    anchor.set_from_marker(from);
    anchor.set_to_marker(to);
    anchor.set_shape(shape);

    anchor
}

// Helper function to parse color
fn parse_color(color: &str) -> RgbColorModelHex {
    let mut rgb = RgbColorModelHex::default();

    match color.to_lowercase().as_str() {
        "transparent" => {
            rgb.set_val("FFFFFF");
            let mut alpha = PositiveFixedPercentageType::default();
            alpha.set_val(0);
            rgb.set_alpha(alpha);
        }
        "red" => {
            rgb.set_val("FF0000");
        }
        "green" => {
            rgb.set_val("00FF00");
        }
        "blue" => {
            rgb.set_val("0000FF");
        }
        "yellow" => {
            rgb.set_val("FFFF00");
        }
        "black" => {
            rgb.set_val("000000");
        }
        "white" => {
            rgb.set_val("FFFFFF");
        }
        "gray" | "grey" => {
            rgb.set_val("808080");
        }
        hex if hex.starts_with("#") => {
            rgb.set_val(&hex[1..]);
        }
        hex => {
            rgb.set_val(hex);
        }
    }

    rgb
}

// Helper function to create a cell marker from a cell address string
fn get_cell_marker(cell_address: &str) -> CellMarker {
    // Create a new cell marker
    let mut marker = CellMarker::default();

    // Extract column and row from address (e.g., "A1" -> col=0, row=0)
    let col_index = umya_spreadsheet::helper::coordinate::column_index_from_string(
        &cell_address
            .chars()
            .take_while(|c| c.is_alphabetic())
            .collect::<String>(),
    );

    let row_index = cell_address
        .chars()
        .skip_while(|c| c.is_alphabetic())
        .collect::<String>()
        .parse::<u32>()
        .unwrap_or(1);

    // Set the marker coordinates (0-based indexing)
    marker.set_col(col_index - 1);
    marker.set_row(row_index - 1);

    marker
}
