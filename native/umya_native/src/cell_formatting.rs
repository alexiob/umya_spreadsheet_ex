use crate::{atoms, UmyaSpreadsheet};
use rustler::{Atom, Error as NifError, NifResult, ResourceArc};
use umya_spreadsheet::{EnumTrait, FontSchemeValues, UnderlineValues};

#[rustler::nif]
fn set_font_italic(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    is_italic: bool,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let cell = sheet.get_cell_mut(&*cell_address);
            cell.get_style_mut().get_font_mut().set_italic(is_italic);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn set_font_underline(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    underline_type: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let cell = sheet.get_cell_mut(&*cell_address);

            // Map underline type string to UnderlineValues enum
            let underline_value = match underline_type.as_str() {
                "none" => UnderlineValues::None,
                "single" => UnderlineValues::Single,
                "double" => UnderlineValues::Double,
                "single_accounting" => UnderlineValues::SingleAccounting,
                "double_accounting" => UnderlineValues::DoubleAccounting,
                _ => UnderlineValues::Single, // Default to single if invalid
            };

            cell.get_style_mut()
                .get_font_mut()
                .set_underline(underline_value.get_value_string());
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn set_font_strikethrough(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    is_strikethrough: bool,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let cell = sheet.get_cell_mut(&*cell_address);
            let font = cell.get_style_mut().get_font_mut();
            font.set_strikethrough(is_strikethrough);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn set_border_style(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    border_position: String,
    border_style: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let cell = sheet.get_cell_mut(&*cell_address);
            let borders = cell.get_style_mut().get_borders_mut();

            match border_position.as_str() {
                "top" => borders.get_top_mut().set_border_style(border_style),
                "bottom" => borders.get_bottom_mut().set_border_style(border_style),
                "left" => borders.get_left_mut().set_border_style(border_style),
                "right" => borders.get_right_mut().set_border_style(border_style),
                "diagonal" => {
                    borders.get_diagonal_mut().set_border_style(border_style);
                    // Enable diagonal up/down by default when setting diagonal border
                    borders.set_diagonal_up(true);
                    borders.set_diagonal_down(true);
                }
                "all" => {
                    // Set all borders to the same style
                    borders.get_top_mut().set_border_style(border_style.clone());
                    borders
                        .get_bottom_mut()
                        .set_border_style(border_style.clone());
                    borders
                        .get_left_mut()
                        .set_border_style(border_style.clone());
                    borders.get_right_mut().set_border_style(border_style);
                }
                _ => {
                    return Err(NifError::Term(Box::new((
                        atoms::error(),
                        "Invalid border position".to_string(),
                    ))))
                }
            }

            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn set_cell_rotation(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    rotation: i32,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let cell = sheet.get_cell_mut(&*cell_address);

            // Convert negative angles to the format Excel expects (0-359)
            let rotation_value = if rotation < 0 {
                (360 + rotation) as u32
            } else {
                rotation as u32
            };

            cell.get_style_mut()
                .get_alignment_mut()
                .set_text_rotation(rotation_value);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn set_cell_indent(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    indent: u32,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // For Excel, indentation is typically implemented as part of the cell format
            // Since direct indent method is not available, we'll have to use a workaround
            // or fall back to a simpler implementation

            // Here we'll set the horizontal alignment to "left" and simulate indentation
            // with cell formatting when possible
            let cell = sheet.get_cell_mut(&*cell_address);
            cell.get_style_mut()
                .get_alignment_mut()
                .set_horizontal(umya_spreadsheet::HorizontalAlignmentValues::Left);

            // Since we can't directly set indentation, we'll set a comment indicating it
            // This is a placeholder until we find a proper implementation
            let cell_value = cell.get_value().to_string();
            if indent > 0 {
                let spaces = " ".repeat(indent as usize * 2); // Use 2 spaces per indent level
                cell.set_value(&format!("{}{}", spaces, cell_value));
            }

            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn set_font_family(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    font_family: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let cell = sheet.get_cell_mut(&*cell_address);

            // Map font family string to numeric values
            // Roman/Serif (1), Swiss/Sans-Serif (2), Modern/Monospace (3)
            // Script (4), Decorative (5)
            let family_value = match font_family.to_lowercase().as_str() {
                "roman" => 1,      // Serif fonts like Times New Roman
                "swiss" => 2,      // Sans-Serif fonts like Arial
                "modern" => 3,     // Monospaced fonts like Courier
                "script" => 4,     // Script fonts
                "decorative" => 5, // Decorative fonts
                _ => 2,            // Default to Swiss/Sans-serif (most common)
            };

            cell.get_style_mut().get_font_mut().set_family(family_value);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn set_font_scheme(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    font_scheme: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let cell = sheet.get_cell_mut(&*cell_address);

            // Map scheme type string to FontSchemeValues enum
            let scheme_value = match font_scheme.to_lowercase().as_str() {
                "major" => FontSchemeValues::Major,
                "minor" => FontSchemeValues::Minor,
                "none" => FontSchemeValues::None,
                _ => FontSchemeValues::None, // Default to none if invalid
            };

            cell.get_style_mut()
                .get_font_mut()
                .set_scheme(scheme_value.get_value_string());
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}
