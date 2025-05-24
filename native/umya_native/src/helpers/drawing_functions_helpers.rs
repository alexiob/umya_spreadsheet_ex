// Helper function to get a mutable reference to the spreadsheet
fn get_spreadsheet_handle(
    spreadsheet: &ResourceArc<Mutex<Spreadsheet>>,
) -> Result<std::sync::MutexGuard<'_, Spreadsheet>, Atom> {
    match spreadsheet.lock() {
        Ok(handle) => Ok(handle),
        Err(_) => Err(error()),
    }
}

// Helper function to get a mutable reference to a worksheet by name
fn get_worksheet_by_name_mut<'a>(
    spreadsheet: &'a mut Spreadsheet,
    sheet_name: &str,
) -> Result<&'a mut umya_spreadsheet::Worksheet, Atom> {
    match spreadsheet.get_sheet_by_name_mut(sheet_name) {
        Some(worksheet) => Ok(worksheet),
        None => Err(atoms::error()),
    }
}

// Helper function to create a cell marker from a cell address string
fn get_cell_marker(cell_address: &str) -> CellMarker {
    // Parse cell address to get column and row
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
    shape_props.set_preset_geometry(preset_geometry);

    // Set fill color
    let rgb_fill = parse_color(fill_color);
    let mut solid_fill = SolidFill::default();
    solid_fill.set_color(rgb_fill);
    shape_props.set_solid_fill(solid_fill);

    // Set outline
    let rgb_outline = parse_color(outline_color);
    let mut outline = Outline::default();
    outline.set_color(rgb_outline);
    outline.set_width(outline_width);
    shape_props.set_outline(outline);

    // Create transformation matrix
    let mut transform = Transform2D::default();
    transform.set_offset(Point2D::new(0.0, 0.0));
    transform.set_extents(PositiveSize2D::new(width, height));
    shape_props.set_transform(transform);

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
    preset_geometry.set_preset("straightConnector1");
    shape_props.set_preset_geometry(preset_geometry);

    // Set line color and width
    let rgb_line = parse_color(line_color);
    let mut outline = Outline::default();
    outline.set_color(rgb_line);
    outline.set_width(line_width);
    shape_props.set_outline(outline);

    // No fill for connector
    let mut solid_fill = SolidFill::default();
    solid_fill.set_color(parse_color("transparent"));
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
fn parse_color(color: &str) -> RGBColorModelHex {
    let mut rgb = RGBColorModelHex::default();

    match color.to_lowercase().as_str() {
        "transparent" => rgb.set_rgb_color("FFFFFF").set_alpha(0),
        "red" => rgb.set_rgb_color("FF0000"),
        "green" => rgb.set_rgb_color("00FF00"),
        "blue" => rgb.set_rgb_color("0000FF"),
        "yellow" => rgb.set_rgb_color("FFFF00"),
        "black" => rgb.set_rgb_color("000000"),
        "white" => rgb.set_rgb_color("FFFFFF"),
        hex if hex.starts_with("#") => rgb.set_rgb_color(&hex[1..]),
        hex => rgb.set_rgb_color(hex),
    }

    rgb
}
