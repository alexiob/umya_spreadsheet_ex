use crate::{atoms, UmyaSpreadsheet};
use rustler::{Error as NifError, NifResult, ResourceArc};
use umya_spreadsheet::EnumTrait;

#[rustler::nif]
pub fn get_font_name(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<String> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            if let Some(cell) = sheet.get_cell(&*cell_address) {
                if let Some(font) = cell.get_style().get_font() {
                    Ok(font.get_name().to_string())
                } else {
                    Ok("Calibri".to_string()) // Default font name
                }
            } else {
                Ok("Calibri".to_string()) // Default font name for non-existing cell
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn get_font_size(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<f64> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            let cell = sheet.get_cell(&*cell_address);
            if let Some(cell) = cell {
                if let Some(font) = cell.get_style().get_font() {
                    Ok(*font.get_size())
                } else {
                    Ok(11.0) // Default font size
                }
            } else {
                Ok(11.0) // Default font size for non-existing cell
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn get_font_bold(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<bool> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            let cell = sheet.get_cell(&*cell_address);
            if let Some(cell) = cell {
                if let Some(font) = cell.get_style().get_font() {
                    Ok(*font.get_bold())
                } else {
                    Ok(false) // Default not bold
                }
            } else {
                Ok(false) // Default not bold for non-existing cell
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn get_font_italic(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<bool> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            let cell = sheet.get_cell(&*cell_address);
            if let Some(cell) = cell {
                if let Some(font) = cell.get_style().get_font() {
                    Ok(*font.get_italic())
                } else {
                    Ok(false) // Default not italic
                }
            } else {
                Ok(false) // Default not italic for non-existing cell
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn get_font_underline(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<String> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            let cell = sheet.get_cell(&*cell_address);
            if let Some(cell) = cell {
                if let Some(font) = cell.get_style().get_font() {
                    Ok(font.get_underline().to_string())
                } else {
                    Ok("none".to_string()) // Default no underline
                }
            } else {
                Ok("none".to_string()) // Default no underline for non-existing cell
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn get_font_strikethrough(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<bool> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            let cell = sheet.get_cell(&*cell_address);
            if let Some(cell) = cell {
                if let Some(font) = cell.get_style().get_font() {
                    Ok(*font.get_strikethrough())
                } else {
                    Ok(false) // Default no strikethrough
                }
            } else {
                Ok(false) // Default no strikethrough for non-existing cell
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn get_font_family(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<String> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            let cell = sheet.get_cell(&*cell_address);
            if let Some(cell) = cell {
                if let Some(font) = cell.get_style().get_font() {
                    let family_value = *font.get_family();

                    // Convert numeric family value back to string representation
                    let family_name = match family_value {
                        1 => "roman",      // Serif fonts like Times New Roman
                        2 => "swiss",      // Sans-Serif fonts like Arial
                        3 => "modern",     // Monospaced fonts like Courier
                        4 => "script",     // Script fonts
                        5 => "decorative", // Decorative fonts
                        _ => "auto",       // Auto/default family type
                    };

                    Ok(family_name.to_string())
                } else {
                    Ok("auto".to_string()) // Default family
                }
            } else {
                Ok("auto".to_string()) // Default family for non-existing cell
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn get_font_scheme(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<String> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            let cell = sheet.get_cell(&*cell_address);
            if let Some(cell) = cell {
                if let Some(font) = cell.get_style().get_font() {
                    Ok(font.get_scheme().to_string())
                } else {
                    Ok("none".to_string()) // Default scheme
                }
            } else {
                Ok("none".to_string()) // Default scheme for non-existing cell
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn get_font_color(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<String> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            let cell = sheet.get_cell(&*cell_address);
            if let Some(cell) = cell {
                if let Some(font) = cell.get_style().get_font() {
                    let color = font.get_color();

                    // Try to get ARGB value, or return default if not set
                    let argb = color.get_argb();
                    if !argb.is_empty() {
                        Ok(argb.to_string())
                    } else {
                        let theme_index = color.get_theme_index();
                        if *theme_index > 0 {
                            Ok(format!("theme:{}", theme_index))
                        } else {
                            Ok("FF000000".to_string()) // Default black color
                        }
                    }
                } else {
                    Ok("FF000000".to_string()) // Default black color
                }
            } else {
                Ok("FF000000".to_string()) // Default black color for non-existing cell
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

// ============================================================================
// ALIGNMENT GETTERS
// ============================================================================

#[rustler::nif]
pub fn get_cell_horizontal_alignment(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<String> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            if let Some(cell) = sheet.get_cell(&*cell_address) {
                if let Some(alignment) = cell.get_style().get_alignment() {
                    Ok(alignment.get_horizontal().get_value_string().to_string())
                } else {
                    Ok("general".to_string()) // Default horizontal alignment
                }
            } else {
                Ok("general".to_string()) // Default for non-existing cell
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn get_cell_vertical_alignment(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<String> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            if let Some(cell) = sheet.get_cell(&*cell_address) {
                if let Some(alignment) = cell.get_style().get_alignment() {
                    let vertical_value = alignment.get_vertical().get_value_string();
                    // Convert internal "center" to external "middle" for consistency
                    let result = if vertical_value == "center" {
                        "middle".to_string()
                    } else {
                        vertical_value.to_string()
                    };
                    Ok(result)
                } else {
                    Ok("bottom".to_string()) // Default vertical alignment
                }
            } else {
                Ok("bottom".to_string()) // Default for non-existing cell
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn get_cell_wrap_text(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<bool> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            if let Some(cell) = sheet.get_cell(&*cell_address) {
                if let Some(alignment) = cell.get_style().get_alignment() {
                    Ok(*alignment.get_wrap_text())
                } else {
                    Ok(false) // Default wrap text value
                }
            } else {
                Ok(false) // Default for non-existing cell
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn get_cell_text_rotation(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<u32> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            if let Some(cell) = sheet.get_cell(&*cell_address) {
                if let Some(alignment) = cell.get_style().get_alignment() {
                    Ok(*alignment.get_text_rotation())
                } else {
                    Ok(0) // Default text rotation
                }
            } else {
                Ok(0) // Default for non-existing cell
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

// ============================================================================
// BORDER GETTERS
// ============================================================================

#[rustler::nif]
pub fn get_border_style(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    border_position: String,
) -> NifResult<String> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            if let Some(cell) = sheet.get_cell(&*cell_address) {
                if let Some(borders) = cell.get_style().get_borders() {
                    let border_style = match border_position.as_str() {
                        "top" => borders.get_top().get_style().get_value_string(),
                        "bottom" => borders.get_bottom().get_style().get_value_string(),
                        "left" => borders.get_left().get_style().get_value_string(),
                        "right" => borders.get_right().get_style().get_value_string(),
                        "diagonal" => borders.get_diagonal().get_style().get_value_string(),
                        _ => "none",
                    };
                    Ok(border_style.to_string())
                } else {
                    Ok("none".to_string()) // Default border style
                }
            } else {
                Ok("none".to_string()) // Default for non-existing cell
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn get_border_color(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    border_position: String,
) -> NifResult<String> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            if let Some(cell) = sheet.get_cell(&*cell_address) {
                if let Some(borders) = cell.get_style().get_borders() {
                    let border_color = match border_position.as_str() {
                        "top" => {
                            let color = borders.get_top().get_color().get_argb();
                            if color.is_empty() {
                                "FF000000"
                            } else {
                                color
                            }
                        }
                        "bottom" => {
                            let color = borders.get_bottom().get_color().get_argb();
                            if color.is_empty() {
                                "FF000000"
                            } else {
                                color
                            }
                        }
                        "left" => {
                            let color = borders.get_left().get_color().get_argb();
                            if color.is_empty() {
                                "FF000000"
                            } else {
                                color
                            }
                        }
                        "right" => {
                            let color = borders.get_right().get_color().get_argb();
                            if color.is_empty() {
                                "FF000000"
                            } else {
                                color
                            }
                        }
                        "diagonal" => {
                            let color = borders.get_diagonal().get_color().get_argb();
                            if color.is_empty() {
                                "FF000000"
                            } else {
                                color
                            }
                        }
                        _ => "FF000000", // Default black
                    };
                    Ok(border_color.to_string())
                } else {
                    Ok("FF000000".to_string()) // Default black border color
                }
            } else {
                Ok("FF000000".to_string()) // Default for non-existing cell
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

// ============================================================================
// FILL/BACKGROUND GETTERS
// ============================================================================

#[rustler::nif]
pub fn get_cell_background_color(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<String> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            if let Some(cell) = sheet.get_cell(&*cell_address) {
                if let Some(fill) = cell.get_style().get_fill() {
                    if let Some(pattern_fill) = fill.get_pattern_fill() {
                        // For solid fills, the "background" color is stored as the foreground color
                        if let Some(fg_color) = pattern_fill.get_foreground_color() {
                            Ok(fg_color.get_argb().to_string())
                        } else {
                            Ok("FFFFFFFF".to_string()) // Default white background
                        }
                    } else {
                        Ok("FFFFFFFF".to_string()) // Default white background
                    }
                } else {
                    Ok("FFFFFFFF".to_string()) // Default white background
                }
            } else {
                Ok("FFFFFFFF".to_string()) // Default for non-existing cell
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn get_cell_foreground_color(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<String> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            if let Some(cell) = sheet.get_cell(&*cell_address) {
                if let Some(fill) = cell.get_style().get_fill() {
                    if let Some(pattern_fill) = fill.get_pattern_fill() {
                        // For pattern fills, the "foreground color" typically refers to
                        // the background color of the pattern, not the visible background
                        if let Some(bg_color) = pattern_fill.get_background_color() {
                            Ok(bg_color.get_argb().to_string())
                        } else {
                            Ok("FF000000".to_string()) // Default black foreground
                        }
                    } else {
                        Ok("FF000000".to_string()) // Default black foreground
                    }
                } else {
                    Ok("FF000000".to_string()) // Default black foreground
                }
            } else {
                Ok("FF000000".to_string()) // Default for non-existing cell
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn get_cell_pattern_type(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<String> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            if let Some(cell) = sheet.get_cell(&*cell_address) {
                if let Some(fill) = cell.get_style().get_fill() {
                    if let Some(pattern_fill) = fill.get_pattern_fill() {
                        Ok(pattern_fill
                            .get_pattern_type()
                            .get_value_string()
                            .to_string())
                    } else {
                        Ok("none".to_string()) // Default pattern type
                    }
                } else {
                    Ok("none".to_string()) // Default pattern type
                }
            } else {
                Ok("none".to_string()) // Default for non-existing cell
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

// ============================================================================
// NUMBER FORMAT GETTERS
// ============================================================================

#[rustler::nif]
pub fn get_cell_number_format_id(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<u32> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            if let Some(cell) = sheet.get_cell(&*cell_address) {
                if let Some(number_format) = cell.get_style().get_numbering_format() {
                    Ok(*number_format.get_number_format_id())
                } else {
                    Ok(0) // Default number format ID (General)
                }
            } else {
                Ok(0) // Default for non-existing cell
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn get_cell_format_code(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<String> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            if let Some(cell) = sheet.get_cell(&*cell_address) {
                if let Some(number_format) = cell.get_style().get_numbering_format() {
                    Ok(number_format.get_format_code().to_string())
                } else {
                    Ok("General".to_string()) // Default format code
                }
            } else {
                Ok("General".to_string()) // Default for non-existing cell
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

// ============================================================================
// PROTECTION GETTERS
// ============================================================================

#[rustler::nif]
pub fn get_cell_locked(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<bool> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            if let Some(cell) = sheet.get_cell(&*cell_address) {
                if let Some(protection) = cell.get_style().get_protection() {
                    Ok(*protection.get_locked())
                } else {
                    Ok(true) // Default locked value
                }
            } else {
                Ok(true) // Default for non-existing cell
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn get_cell_hidden(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    _cell_address: String,
) -> NifResult<bool> {
    let _guard = resource.spreadsheet.lock().unwrap();

    match _guard.get_sheet_by_name(&sheet_name) {
        Some(_sheet) => {
            // Note: The Rust library's Protection::get_hidden() requires a mutable reference
            // which we cannot obtain in a getter function. This needs to be fixed in the
            // underlying Rust library. For now, return the default value.
            Ok(false) // Default hidden value
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}
