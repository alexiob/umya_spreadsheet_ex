use crate::{atoms, UmyaSpreadsheet};
use rustler::{Error as NifError, NifResult, ResourceArc};

#[rustler::nif]
fn get_font_name(
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
fn get_font_size(
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
fn get_font_bold(
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
fn get_font_italic(
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
fn get_font_underline(
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
fn get_font_strikethrough(
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
fn get_font_family(
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
fn get_font_scheme(
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
fn get_font_color(
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
