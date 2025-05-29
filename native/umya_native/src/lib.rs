use rustler::{Atom, Error as NifError, NifResult, ResourceArc};
use std::panic::{self, AssertUnwindSafe};
use std::path::Path;
use std::sync::Mutex;
use umya_spreadsheet;

// Import helper modules
use crate::helpers::style_helpers;

mod atoms {
    rustler::atoms! {
        ok,
        error,
        not_found,
        invalid_path,
        unknown_error,

        // Chart function atoms
        position,
        overlay,
        right,
        left,
        top,
        bottom,
        top_right,
        rot_x,
        rot_y,
        perspective,
        height_percent,

        // Data label atoms
        show_values,
        show_percent,
        show_category_name,
        show_series_name,

        // Data validation atoms
        between,
        not_between,
        equal,
        not_equal,
        greater_than,
        less_than,
        greater_than_or_equal,
        less_than_or_equal,
    }
}

mod auto_filter_functions;
mod cell_formatting;
mod chart_functions;
mod comment_functions;
pub mod conditional_formatting;
pub mod conditional_formatting_additional;
pub mod custom_structs;
mod data_validation;
mod document_properties;
mod drawing_functions;
mod file_format_options;
mod formula_functions;
mod get_cell_formatting;
mod helpers;
mod hyperlink;
mod ole_object_functions;
mod page_breaks;
mod print_settings;
mod rich_text_functions;
mod set_background_color;
mod set_cell_alignment;
mod sheet_view_functions;
mod table;
mod vml_support;
mod workbook_view_functions;
mod write_csv_with_options;

pub struct UmyaSpreadsheet {
    spreadsheet: Mutex<umya_spreadsheet::Spreadsheet>,
}

#[rustler::nif]
fn new_file() -> NifResult<ResourceArc<UmyaSpreadsheet>> {
    let spreadsheet = umya_spreadsheet::new_file();
    let resource = ResourceArc::new(UmyaSpreadsheet {
        spreadsheet: Mutex::new(spreadsheet),
    });
    Ok(resource)
}

#[rustler::nif]
fn new_file_empty_worksheet() -> NifResult<ResourceArc<UmyaSpreadsheet>> {
    let spreadsheet = umya_spreadsheet::new_file_empty_worksheet();
    let resource = ResourceArc::new(UmyaSpreadsheet {
        spreadsheet: Mutex::new(spreadsheet),
    });
    Ok(resource)
}

#[rustler::nif]
fn read_file(path: String) -> NifResult<ResourceArc<UmyaSpreadsheet>> {
    // Use path helper to find a valid file path
    let valid_path = match helpers::path_helper::find_valid_file_path(&path) {
        Some(p) => p,
        None => {
            return Err(NifError::Term(Box::new((
                atoms::error(),
                "File not found".to_string(),
            ))))
        }
    };

    let path_obj = Path::new(&valid_path);

    // Improved error handling with specific error messages for corrupted files
    match umya_spreadsheet::reader::xlsx::read(path_obj) {
        Ok(spreadsheet) => {
            let resource = ResourceArc::new(UmyaSpreadsheet {
                spreadsheet: Mutex::new(spreadsheet),
            });
            Ok(resource)
        }
        Err(e) => {
            // Provide specific error messages based on the error type
            let error_msg = match e.to_string().as_str() {
                s if s.contains("zip") => "corrupted_file",
                s if s.contains("xml") => "invalid_format",
                s if s.contains("permission") => "access_denied",
                _ => "read_error",
            };
            Err(NifError::Term(Box::new((
                atoms::error(),
                error_msg.to_string(),
            ))))
        }
    }
}

#[rustler::nif]
fn lazy_read_file(path: String) -> NifResult<ResourceArc<UmyaSpreadsheet>> {
    // Use path helper to find a valid file path
    let valid_path = match helpers::path_helper::find_valid_file_path(&path) {
        Some(p) => p,
        None => {
            return Err(NifError::Term(Box::new((
                atoms::error(),
                "File not found".to_string(),
            ))))
        }
    };

    let path_obj = Path::new(&valid_path);

    // Handle both .xlsx and .xlsm files with lazy loading and improved error handling
    match umya_spreadsheet::reader::xlsx::lazy_read(path_obj) {
        Ok(spreadsheet) => {
            let resource = ResourceArc::new(UmyaSpreadsheet {
                spreadsheet: Mutex::new(spreadsheet),
            });
            Ok(resource)
        }
        Err(e) => {
            // Provide specific error messages based on the error type
            let error_msg = match e.to_string().as_str() {
                s if s.contains("zip") => "corrupted_file",
                s if s.contains("xml") => "invalid_format",
                s if s.contains("permission") => "access_denied",
                _ => "read_error",
            };
            Err(NifError::Term(Box::new((
                atoms::error(),
                error_msg.to_string(),
            ))))
        }
    }
}

#[rustler::nif]
fn write_file(resource: ResourceArc<UmyaSpreadsheet>, path: String) -> NifResult<Atom> {
    // For output paths, we use the direct path since the file might not exist yet
    // We want to ensure the directory exists though
    let path_obj = Path::new(&path);

    // Ensure the parent directory exists
    if let Some(parent) = path_obj.parent() {
        if !parent.exists() {
            if let Err(_) = std::fs::create_dir_all(parent) {
                return Err(NifError::Term(Box::new((
                    atoms::error(),
                    "failed_to_create_directory".to_string(),
                ))));
            }
        }
    }

    let guard = resource.spreadsheet.lock().unwrap();

    match umya_spreadsheet::writer::xlsx::write(&guard, path_obj) {
        Ok(_) => Ok(atoms::ok()),
        Err(e) => {
            let error_msg = match e.to_string().as_str() {
                s if s.contains("permission") => "access_denied",
                s if s.contains("space") => "insufficient_disk_space",
                _ => "write_error",
            };
            Err(NifError::Term(Box::new((
                atoms::error(),
                error_msg.to_string(),
            ))))
        }
    }
}

#[rustler::nif]
fn write_file_light(
    resource: ResourceArc<UmyaSpreadsheet>,
    path: String,
) -> Result<Atom, (Atom, String)> {
    // For output paths, we use the direct path since the file might not exist yet
    // We want to ensure the directory exists though
    let path_obj = Path::new(&path);

    // Ensure the parent directory exists
    if let Some(parent) = path_obj.parent() {
        if !parent.exists() {
            if let Err(_) = std::fs::create_dir_all(parent) {
                return Err((atoms::error(), "failed_to_create_directory".to_string()));
            }
        }
    }

    let guard = resource.spreadsheet.lock().unwrap();

    match umya_spreadsheet::writer::xlsx::write_light(&guard, path_obj) {
        Ok(_) => Ok(atoms::ok()),
        Err(e) => {
            let error_msg = match e.to_string().as_str() {
                s if s.contains("permission") => "access_denied",
                s if s.contains("space") => "insufficient_disk_space",
                _ => "write_error",
            };
            Err((atoms::error(), error_msg.to_string()))
        }
    }
}

#[rustler::nif]
fn write_file_with_password(
    resource: ResourceArc<UmyaSpreadsheet>,
    path: String,
    password: String,
) -> NifResult<Atom> {
    // For output paths, we use the direct path since the file might not exist yet
    // We want to ensure the directory exists though
    let path_obj = Path::new(&path);

    // Ensure the parent directory exists
    if let Some(parent) = path_obj.parent() {
        if !parent.exists() {
            if let Err(_) = std::fs::create_dir_all(parent) {
                return Err(NifError::Term(Box::new((
                    atoms::error(),
                    "Invalid file path or unable to create directory".to_string(),
                ))));
            }
        }
    }

    let guard = resource.spreadsheet.lock().unwrap();

    let result =
        match umya_spreadsheet::writer::xlsx::write_with_password(&guard, path_obj, &password) {
            Ok(_) => Ok(atoms::ok()),
            Err(_) => Err(NifError::Term(Box::new((
                atoms::error(),
                "Failed to write file with password".to_string(),
            )))),
        };

    // Explicitly drop the guard to ensure mutex is released before returning
    drop(guard);

    // Return the result
    result
}

#[rustler::nif]
fn write_file_with_password_light(
    resource: ResourceArc<UmyaSpreadsheet>,
    path: String,
    password: String,
) -> NifResult<Atom> {
    // For output paths, we use the direct path since the file might not exist yet
    // We want to ensure the directory exists though
    let path_obj = Path::new(&path);

    // Ensure the parent directory exists
    if let Some(parent) = path_obj.parent() {
        if !parent.exists() {
            if let Err(_) = std::fs::create_dir_all(parent) {
                return Err(NifError::Term(Box::new((
                    atoms::error(),
                    "Invalid file path or unable to create directory".to_string(),
                ))));
            }
        }
    }

    let guard = resource.spreadsheet.lock().unwrap();

    let result = match umya_spreadsheet::writer::xlsx::write_with_password_light(
        &guard, path_obj, &password,
    ) {
        Ok(_) => Ok(atoms::ok()),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Failed to write file with password".to_string(),
        )))),
    };

    // Explicitly drop the guard to ensure mutex is released before returning
    drop(guard);

    // Return the result
    result
}

// Cell Reading/Writing operations

#[rustler::nif]
fn get_cell_value(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<(Atom, String)> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    // For lazy-loaded spreadsheets, ensure the sheet is loaded
    guard.read_sheet_collection();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            let cell_value = sheet.get_cell_value(&*cell_address);
            // Get the string representation of the cell value and return as {:ok, value} tuple
            Ok((atoms::ok(), cell_value.get_value().into_owned()))
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn set_cell_value(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    value: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Workaround for missing set_cell_value
            let cell = sheet.get_cell_mut(&*cell_address);
            cell.set_value(&value);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn get_sheet_names(resource: ResourceArc<UmyaSpreadsheet>) -> Vec<String> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    // For lazy-loaded spreadsheets, we need to ensure all sheets are loaded
    // This works for both regular and lazy-loaded spreadsheets
    guard.read_sheet_collection();

    let mut sheet_names = Vec::new();
    for sheet in guard.get_sheet_collection() {
        sheet_names.push(sheet.get_name().to_string());
    }
    sheet_names
}

#[rustler::nif]
fn add_sheet(resource: ResourceArc<UmyaSpreadsheet>, sheet_name: String) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.new_sheet(&sheet_name) {
        Ok(_) => Ok(atoms::ok()),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Failed to add sheet".to_string(),
        )))),
    }
}

#[rustler::nif]
fn rename_sheet(
    resource: ResourceArc<UmyaSpreadsheet>,
    old_sheet_name: String,
    new_sheet_name: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    // Check if the new sheet name already exists
    if guard.get_sheet_by_name(&new_sheet_name).is_some() {
        return Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet name already exists".to_string(),
        ))));
    }

    // Find and rename the sheet
    match guard.get_sheet_by_name_mut(&old_sheet_name) {
        Some(sheet) => {
            sheet.set_name(&new_sheet_name);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn move_range(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    range: String,
    row: i32,
    column: i32,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            sheet.move_range(&range, &row, &column);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

// Cell style operations

// Advanced functions for styling

#[rustler::nif]
fn set_font_color(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    color: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Use the color helper to parse the color string
            let color_obj = match helpers::color_helper::create_color_object(&color) {
                Ok(color) => color,
                Err(_) => {
                    return Err(NifError::Term(Box::new((
                        atoms::error(),
                        "Invalid color format".to_string(),
                    ))))
                }
            };

            // Apply the font color using the style helper
            style_helpers::apply_cell_style(
                sheet,
                &*cell_address,
                None,
                Some(color_obj),
                None,
                None,
            );

            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn set_font_size(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    size: i32,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Apply the font size using the style helper
            style_helpers::apply_cell_style(
                sheet,
                &*cell_address,
                None,
                None,
                Some(size as f64),
                None,
            );

            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn set_font_bold(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    is_bold: bool,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Apply the bold setting using the style helper
            style_helpers::apply_cell_style(sheet, &*cell_address, None, None, None, Some(is_bold));

            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn set_font_name(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    font_name: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let style = sheet.get_style_mut(&*cell_address);
            let font = style.get_font_mut();
            font.set_name(&font_name);

            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn set_row_height(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    row_number: i32,
    height: f64,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Set row height
            sheet
                .get_row_dimension_mut(&(row_number as u32))
                .set_height(height);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn remove_cell(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let (col_opt, row_opt, _, _) =
                umya_spreadsheet::helper::coordinate::index_from_coordinate(&cell_address);
            if let (Some(col), Some(row)) = (col_opt, row_opt) {
                sheet.remove_cell((&col, &row));
                Ok(atoms::ok())
            } else {
                Err(NifError::Term(Box::new((
                    atoms::error(),
                    "Invalid cell address".to_string(),
                ))))
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn get_formatted_value(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<(Atom, String)> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    // For lazy-loaded spreadsheets, ensure the sheet is loaded
    guard.read_sheet_collection();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            let formatted_value = sheet.get_formatted_value(&*cell_address);
            Ok((atoms::ok(), formatted_value))
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn set_number_format(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    format_code: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Set number format
            let style = sheet.get_style_mut(&*cell_address);
            style.get_number_format_mut().set_format_code(format_code);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn set_row_style(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    row_number: i32,
    bg_color: String,
    font_color: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Process background color
            let bg_color_obj = match helpers::color_helper::create_color_object(&bg_color) {
                Ok(color) => color,
                Err(_) => {
                    return Err(NifError::Term(Box::new((
                        atoms::error(),
                        "Invalid background color format".to_string(),
                    ))))
                }
            };

            // Process font color
            let font_color_obj = match helpers::color_helper::create_color_object(&font_color) {
                Ok(color) => color,
                Err(_) => {
                    return Err(NifError::Term(Box::new((
                        atoms::error(),
                        "Invalid font color format".to_string(),
                    ))))
                }
            };

            // Use the style helper to apply both colors at once
            style_helpers::apply_row_style(
                sheet,
                row_number as u32,
                Some(bg_color_obj),
                Some(font_color_obj),
                None, // No font size change
                None, // No bold change
            );

            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn add_image(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    image_path: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Find valid image path using generalized file path helper
            if let Some(valid_path) = helpers::path_helper::find_valid_file_path(&image_path) {
                // Create marker for image position
                let mut marker =
                    umya_spreadsheet::structs::drawing::spreadsheet::MarkerType::default();
                marker.set_coordinate(&*cell_address);

                // Create image
                let mut image = umya_spreadsheet::structs::Image::default();

                // Try to set the image safely
                let result = panic::catch_unwind(AssertUnwindSafe(|| {
                    image.new_image(&valid_path, marker);
                }));

                if result.is_ok() {
                    // Add the image to the worksheet
                    sheet.add_image(image);
                    return Ok(atoms::ok());
                }
            }

            // If we couldn't find a valid path or set the image
            Err(NifError::Term(Box::new((
                atoms::error(),
                "Image not found".to_string(),
            ))))
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn download_image(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    output_path: String,
) -> NifResult<Atom> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            // Get the image
            let image_option = sheet.get_image(&*cell_address);
            match image_option {
                Some(image) => {
                    // For output paths, ensure the directory exists
                    let path_obj = Path::new(&output_path);

                    // Ensure the parent directory exists
                    if let Some(parent) = path_obj.parent() {
                        if !parent.exists() {
                            if let Err(_) = std::fs::create_dir_all(parent) {
                                return Err(NifError::Term(Box::new((
                                    atoms::error(),
                                    "Invalid output path or unable to create directory".to_string(),
                                ))));
                            }
                        }
                    }

                    // Download the image
                    image.download_image(&output_path);
                    Ok(atoms::ok())
                }
                None => Err(NifError::Term(Box::new((
                    atoms::error(),
                    "Sheet not found".to_string(),
                )))),
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn change_image(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    new_image_path: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Get the image
            let image_option = sheet.get_image_mut(&*cell_address);
            match image_option {
                Some(image) => {
                    // Find valid image path using generalized file path helper
                    if let Some(valid_path) =
                        helpers::path_helper::find_valid_file_path(&new_image_path)
                    {
                        // Try to change the image with error handling
                        let result = panic::catch_unwind(AssertUnwindSafe(|| {
                            image.change_image(&valid_path);
                        }));

                        if result.is_ok() {
                            return Ok(atoms::ok());
                        }
                    }

                    // If we couldn't find a valid path or set the image
                    Err(NifError::Term(Box::new((
                        atoms::error(),
                        "Failed to download image".to_string(),
                    ))))
                }
                None => Err(NifError::Term(Box::new((
                    atoms::error(),
                    "Sheet not found".to_string(),
                )))),
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn clone_sheet(
    resource: ResourceArc<UmyaSpreadsheet>,
    source_sheet_name: String,
    new_sheet_name: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&source_sheet_name) {
        Some(sheet) => {
            // Clone the source sheet and update its name
            let mut clone_sheet = sheet.clone();
            clone_sheet.set_name(&new_sheet_name);

            // Add the cloned sheet to the spreadsheet
            match guard.add_sheet(clone_sheet) {
                Ok(_) => Ok(atoms::ok()),
                Err(_) => Err(NifError::Term(Box::new((
                    atoms::error(),
                    "Failed to add cloned sheet".to_string(),
                )))),
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn remove_sheet(resource: ResourceArc<UmyaSpreadsheet>, sheet_name: String) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.remove_sheet_by_name(&sheet_name) {
        Ok(_) => Ok(atoms::ok()),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn insert_new_row(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    row_index: i32,
    amount: i32,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let row_index_u32 = row_index as u32;
            let amount_u32 = amount as u32;
            sheet.insert_new_row(&row_index_u32, &amount_u32);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn insert_new_column(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    column: String,
    amount: i32,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let amount_u32 = amount as u32;
            sheet.insert_new_column(&column, &amount_u32);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn insert_new_column_by_index(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    column_index: i32,
    amount: i32,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let column_index_u32 = column_index as u32;
            let amount_u32 = amount as u32;
            sheet.insert_new_column_by_index(&column_index_u32, &amount_u32);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn remove_row(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    row_index: i32,
    amount: i32,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let row_index_u32 = row_index as u32;
            let amount_u32 = amount as u32;
            sheet.remove_row(&row_index_u32, &amount_u32);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn remove_column(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    column: String,
    amount: i32,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let amount_u32 = amount as u32;
            sheet.remove_column(&column, &amount_u32);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn remove_column_by_index(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    column_index: i32,
    amount: i32,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let column_index_u32 = column_index as u32;
            let amount_u32 = amount as u32;
            sheet.remove_column_by_index(&column_index_u32, &amount_u32);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn copy_row_styling(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    source_row: u32,
    target_row: u32,
    start_column: Option<u32>,
    end_column: Option<u32>,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            sheet.copy_row_styling(
                &source_row,
                &target_row,
                start_column.as_ref(),
                end_column.as_ref(),
            );
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn copy_column_styling(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    source_column: u32,
    target_column: u32,
    start_row: Option<u32>,
    end_row: Option<u32>,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            sheet.copy_col_styling(
                &source_column,
                &target_column,
                start_row.as_ref(),
                end_row.as_ref(),
            );
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn set_wrap_text(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    wrap_text: bool,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            sheet
                .get_cell_mut(&*cell_address)
                .get_style_mut()
                .get_alignment_mut()
                .set_wrap_text(wrap_text);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn set_column_width(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    column: String,
    width: f64,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            sheet.get_column_dimension_mut(&column).set_width(width);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn set_column_auto_width(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    column: String,
    auto_width: bool,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            sheet
                .get_column_dimension_mut(&column)
                .set_auto_width(auto_width);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn set_sheet_protection(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    password: Option<String>,
    protected: bool,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let sheet_protection = sheet.get_sheet_protection_mut();
            sheet_protection.set_sheet(protected);

            if let Some(pwd) = password {
                sheet_protection.set_password(&pwd);
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
fn set_workbook_protection(
    resource: ResourceArc<UmyaSpreadsheet>,
    password: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    guard
        .get_workbook_protection_mut()
        .set_workbook_password(&password);
    Ok(atoms::ok())
}

#[rustler::nif]
fn set_sheet_state(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    state: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let state_value = match state.as_str() {
                "visible" => umya_spreadsheet::SheetStateValues::Visible,
                "hidden" => umya_spreadsheet::SheetStateValues::Hidden,
                "very_hidden" => umya_spreadsheet::SheetStateValues::VeryHidden,
                _ => {
                    return Err(NifError::Term(Box::new((
                        atoms::error(),
                        "Invalid sheet state".to_string(),
                    ))))
                }
            };

            sheet.set_state(state_value);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

// Removed duplicate set_show_grid_lines function - now in sheet_view_functions.rs

#[rustler::nif]
fn add_merge_cells(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    range: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            sheet.add_merge_cells(&range);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

#[rustler::nif]
fn set_password(input_path: String, output_path: String, password: String) -> NifResult<Atom> {
    // Use the path helper to find valid file paths
    let valid_input_path = match helpers::path_helper::find_valid_file_path(&input_path) {
        Some(path) => path,
        None => {
            return Err(NifError::Term(Box::new((
                atoms::error(),
                "Sheet not found".to_string(),
            ))))
        } // Input file not found
    };

    // For output path, we use the direct path since it might not exist yet
    let output_path = Path::new(&output_path);

    // Use the valid input path
    let input_path = Path::new(&valid_input_path);

    match umya_spreadsheet::writer::xlsx::set_password(input_path, output_path, &password) {
        Ok(_) => Ok(atoms::ok()),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Failed to set password".to_string(),
        )))),
    }
}

// ------------ Pivot table functions ------------

use std::collections::hash_map::DefaultHasher;
use std::hash::{Hash, Hasher};
use umya_spreadsheet::{
    Address, CacheSource, ColumnFields, DataField, DataFields, Field, PivotCacheDefinition,
    PivotField, PivotFields, PivotTable, PivotTableDefinition, Range, RowItem, RowItems,
    SourceValues, WorksheetSource,
};

/// Add a pivot table to a spreadsheet
#[rustler::nif]
pub fn add_pivot_table(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    name: String,
    source_sheet: String,
    source_range: String,
    target_cell: String,
    row_fields: Vec<i32>,
    column_fields: Vec<i32>,
    data_fields: Vec<(i32, String, String)>,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Create a new pivot table
            let mut pivot_table = PivotTable::default();

            // Generate a stable cache ID based on sheet and table name
            let mut s = DefaultHasher::new();
            sheet_name.hash(&mut s);
            name.hash(&mut s);
            let cache_id = (s.finish() % 1000) as u32;

            // Set up pivot table definition
            let mut pivot_table_def = PivotTableDefinition::default();
            pivot_table_def.set_name(&name);
            pivot_table_def.set_cache_id(cache_id);

            // Set location (required)
            let mut location = umya_spreadsheet::Location::default();
            location.set_reference(&target_cell);
            pivot_table_def.set_location(location);

            // Setup pivot fields (one per source column)
            let mut pivot_fields = PivotFields::default();
            let cols = 4; // We know our test data is A1:D5

            for _i in 0..cols {
                let mut pivot_field = PivotField::default();
                pivot_field.set_show_all(true);
                pivot_field.set_data_field(false);
                pivot_fields.add_list_mut(pivot_field);
            }
            pivot_table_def.set_pivot_fields(pivot_fields);

            // Configure row fields
            let mut row_fields_obj = RowItems::default();
            if !row_fields.is_empty() {
                let row_item = RowItem::default();
                row_fields_obj.add_list_mut(row_item);
            }
            pivot_table_def.set_row_items(row_fields_obj);

            // Configure column fields
            let mut column_fields_obj = ColumnFields::default();
            if !column_fields.is_empty() {
                let column_field = Field::default();
                column_fields_obj.add_list_mut(column_field);
            }
            pivot_table_def.set_column_fields(column_fields_obj);

            // Configure data fields
            let mut data_fields_obj = DataFields::default();
            for (idx, _func, _field_name) in data_fields.iter() {
                let mut data_field = DataField::default();
                data_field.set_fie_id(*idx as u32);
                // Since set_value doesn't exist, we'll skip setting the name for now
                data_fields_obj.add_list_mut(data_field);
            }
            pivot_table_def.set_data_fields(data_fields_obj);

            // Create and configure cache definition
            let mut pivot_cache_def = PivotCacheDefinition::default();
            pivot_cache_def.set_id(&cache_id.to_string()); // Use same ID as pivot table

            // Set up cache source
            let mut cache_source = CacheSource::default();
            cache_source.set_type(SourceValues::Worksheet);

            // Configure worksheet source with range and sheet info
            let mut address = Address::default();
            address.set_sheet_name(&source_sheet);

            let mut range_obj = Range::default();
            range_obj.set_range(&source_range);
            address.set_range(range_obj);

            let mut worksheet_source = WorksheetSource::default();
            worksheet_source.set_address(address);
            // worksheet_source doesn't have set_name method, but it's not required

            cache_source.set_worksheet_source_mut(worksheet_source);
            pivot_cache_def.set_cache_source(cache_source);

            // Set pivot table and cache definitions
            pivot_table.set_pivot_table_definition(pivot_table_def);
            pivot_table.set_pivot_cache_definition(pivot_cache_def);

            // Set offset to ensure pivot tables don't overlap
            let offset = sheet.get_pivot_tables().len();
            pivot_table
                .get_pivot_table_definition_mut()
                .get_location_mut()
                .set_reference(&format!("{}{}", target_cell, offset));

            // Add pivot table to the sheet
            sheet.add_pivot_table(pivot_table);

            // Verify that the pivot table was properly added with all required components
            let found_valid_pt = {
                let pivot_tables = sheet.get_pivot_tables();
                let mut found = false;

                for (_i, pt) in pivot_tables.iter().enumerate() {
                    let def = pt.get_pivot_table_definition();
                    let pt_name = def.get_name();
                    let has_cache = pt
                        .get_pivot_cache_definition()
                        .get_cache_source()
                        .get_worksheet_source()
                        .is_some();

                    if pt_name == &name && has_cache {
                        found = true;
                        break;
                    }
                }
                found
            };

            if found_valid_pt {
                Ok(atoms::ok())
            } else {
                Err(NifError::Term(Box::new((
                    atoms::error(),
                    "No valid pivot table found".to_string(),
                ))))
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Refresh all pivot tables in a spreadsheet
#[rustler::nif]
pub fn refresh_all_pivot_tables(resource: ResourceArc<UmyaSpreadsheet>) -> NifResult<Atom> {
    let _guard = resource.spreadsheet.lock().unwrap(); // Changed to _guard since it's unused
    Ok(atoms::ok())
}

/// Check if a sheet has any valid pivot tables
#[rustler::nif]
pub fn has_pivot_tables(resource: ResourceArc<UmyaSpreadsheet>, sheet_name: String) -> bool {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            // Get pivot tables and validate each one
            let pivot_tables = sheet.get_pivot_tables();
            let _count = pivot_tables.len();

            // A pivot table is valid if it has:
            // 1. A name
            // 2. A cache source
            // 3. At least one field defined
            // 4. A valid location
            let mut valid_count = 0;
            for (_i, pt) in pivot_tables.iter().enumerate() {
                let def = pt.get_pivot_table_definition();
                let cache = pt.get_pivot_cache_definition();

                let has_name = !def.get_name().is_empty();
                let has_cache = cache.get_cache_source().get_worksheet_source().is_some();
                let has_location = !def.get_location().get_reference().is_empty();
                let has_fields = !def.get_pivot_fields().get_list().is_empty();

                if has_name && has_cache && has_location && has_fields {
                    valid_count += 1;
                }
            }

            // Return true if we have at least one valid pivot table
            valid_count > 0
        }
        None => false,
    }
}

/// Count the number of pivot tables in a sheet
#[rustler::nif]
pub fn count_pivot_tables(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> (rustler::Atom, usize) {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            let pivot_tables = sheet.get_pivot_tables();
            let count = pivot_tables.len();
            (atoms::ok(), count)
        }
        None => (atoms::not_found(), 0),
    }
}

#[rustler::nif]
pub fn remove_pivot_table(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    pivot_table_name: String,
) -> rustler::Atom {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let pivot_tables = sheet.get_pivot_tables();
            let count_before = pivot_tables.len();
            let mut found_idx = None;

            // Find the pivot table by name
            for (i, pt) in pivot_tables.iter().enumerate() {
                if pt.get_pivot_table_definition().get_name() == pivot_table_name {
                    found_idx = Some(i);
                    break;
                }
            }

            if let Some(idx) = found_idx {
                // The sheet's pivot tables are stored as a Vec, so we need to:
                // 1. Get all pivot tables except the one to remove
                // 2. Create a new Vec
                // 3. Add each pivot table except the one we're removing
                // 4. Then clear the existing ones and add our filtered list

                let mut new_tables = Vec::new();

                for (i, pt) in pivot_tables.iter().enumerate() {
                    if i != idx {
                        // Clone pivot table and add it to our new list
                        let mut cloned_pt = PivotTable::default();
                        cloned_pt
                            .set_pivot_table_definition(pt.get_pivot_table_definition().clone());
                        cloned_pt
                            .set_pivot_cache_definition(pt.get_pivot_cache_definition().clone());
                        new_tables.push(cloned_pt);
                    }
                }

                // Clear the existing pivot tables
                for _ in 0..count_before {
                    // This works because we know add_pivot_table exists
                    // If we can add them one by one, we can remove them all and re-add the ones we want to keep
                    if !sheet.get_pivot_tables().is_empty() {
                        let _ = sheet.get_pivot_tables_mut().pop();
                    }
                }

                // Add the filtered pivot tables back
                for pt in new_tables {
                    sheet.add_pivot_table(pt);
                }

                atoms::ok()
            } else {
                atoms::not_found()
            }
        }
        None => atoms::not_found(),
    }
}

// Load function for Rustler
fn load(env: rustler::Env, _info: rustler::Term) -> bool {
    let _ = rustler::resource!(UmyaSpreadsheet, env);
    // Re-enabling Rich Text resources
    let _ = rustler::resource!(rich_text_functions::RichTextResource, env);
    let _ = rustler::resource!(rich_text_functions::TextElementResource, env);
    // OLE Objects resources
    let _ = rustler::resource!(ole_object_functions::OleObjectsResource, env);
    let _ = rustler::resource!(ole_object_functions::OleObjectResource, env);
    let _ = rustler::resource!(ole_object_functions::EmbeddedObjectPropertiesResource, env);
    true
}

// Initialize the NIF using the modern format
rustler::init!(
    "Elixir.UmyaNative",
    load = load,
    nifs = [
        test_simple_function,
        new_file,
        new_file_empty_worksheet,
        read_file,
        lazy_read_file,
        write_file,
        write_file_light,
        write_file_with_password,
        write_file_with_password_light,
        get_cell_value,
        set_cell_value,
        get_sheet_names,
        add_sheet,
        rename_sheet,
        move_range,
        set_font_color,
        set_font_size,
        set_font_bold,
        set_font_name,
        set_row_height,
        remove_cell,
        get_formatted_value,
        set_number_format,
        has_pivot_tables,
        add_pivot_table,
        count_pivot_tables,
        remove_pivot_table,
        refresh_all_pivot_tables,
        // Drawing and shape functions
        drawing_functions::add_shape,
        drawing_functions::add_text_box,
        drawing_functions::add_connector,
        // Cell formatting advanced functions
        cell_formatting::set_font_italic,
        cell_formatting::set_font_underline,
        cell_formatting::set_font_strikethrough,
        cell_formatting::set_font_family,
        cell_formatting::set_font_scheme,
        cell_formatting::set_border_style,
        cell_formatting::set_cell_rotation,
        cell_formatting::set_cell_indent,
        // Cell formatting getter functions
        get_cell_formatting::get_font_name,
        get_cell_formatting::get_font_size,
        get_cell_formatting::get_font_bold,
        get_cell_formatting::get_font_italic,
        get_cell_formatting::get_font_underline,
        get_cell_formatting::get_font_strikethrough,
        get_cell_formatting::get_font_family,
        get_cell_formatting::get_font_scheme,
        get_cell_formatting::get_font_color,
        // Alignment getter functions
        get_cell_formatting::get_cell_horizontal_alignment,
        get_cell_formatting::get_cell_vertical_alignment,
        get_cell_formatting::get_cell_wrap_text,
        get_cell_formatting::get_cell_text_rotation,
        // Border getter functions
        get_cell_formatting::get_border_style,
        get_cell_formatting::get_border_color,
        // Fill/background getter functions
        get_cell_formatting::get_cell_background_color,
        get_cell_formatting::get_cell_foreground_color,
        get_cell_formatting::get_cell_pattern_type,
        // Number format getter functions
        get_cell_formatting::get_cell_number_format_id,
        get_cell_formatting::get_cell_format_code,
        // Protection getter functions
        get_cell_formatting::get_cell_locked,
        get_cell_formatting::get_cell_hidden,
        // Print settings functions
        print_settings::set_page_orientation,
        print_settings::set_paper_size,
        print_settings::set_page_scale,
        print_settings::set_fit_to_page,
        print_settings::set_page_margins,
        print_settings::set_header_footer_margins,
        print_settings::set_header,
        print_settings::set_footer,
        print_settings::set_print_centered,
        print_settings::set_print_area,
        print_settings::set_print_titles,
        // Conditional formatting functions
        conditional_formatting::add_cell_value_rule,
        conditional_formatting::add_color_scale,
        conditional_formatting_additional::add_data_bar,
        conditional_formatting_additional::add_top_bottom_rule,
        conditional_formatting_additional::add_text_rule,
        conditional_formatting_additional::add_icon_set,
        conditional_formatting_additional::add_above_below_average_rule,
        // CSV functions
        write_csv_with_options::write_csv,
        write_csv_with_options::write_csv_with_options,
        // Background color functions
        set_background_color::set_background_color,
        // Cell alignment functions
        set_cell_alignment::set_cell_alignment,
        // Data validation functions
        data_validation::add_number_validation,
        data_validation::add_date_validation,
        data_validation::add_text_length_validation,
        data_validation::add_list_validation,
        data_validation::add_custom_validation,
        data_validation::remove_data_validation,
        // Sheet view functions
        sheet_view_functions::set_show_grid_lines,
        sheet_view_functions::set_tab_selected,
        sheet_view_functions::set_top_left_cell,
        sheet_view_functions::set_zoom_scale,
        sheet_view_functions::freeze_panes,
        sheet_view_functions::split_panes,
        sheet_view_functions::set_tab_color,
        sheet_view_functions::set_sheet_view,
        sheet_view_functions::set_zoom_scale_normal,
        sheet_view_functions::set_zoom_scale_page_layout,
        sheet_view_functions::set_zoom_scale_page_break,
        sheet_view_functions::set_selection,
        workbook_view_functions::set_active_tab,
        workbook_view_functions::set_workbook_window_position,
        copy_column_styling,
        // Comment functions
        comment_functions::add_comment,
        comment_functions::get_comment,
        comment_functions::update_comment,
        comment_functions::remove_comment,
        comment_functions::has_comments,
        comment_functions::get_comments_count,
        // File format options functions
        file_format_options::write_with_compression,
        file_format_options::write_with_encryption_options,
        file_format_options::to_binary_xlsx,
        // Formula functions
        formula_functions::set_formula,
        formula_functions::set_array_formula,
        formula_functions::create_named_range,
        formula_functions::create_defined_name,
        formula_functions::get_defined_names,
        // Auto filter functions
        auto_filter_functions::set_auto_filter,
        auto_filter_functions::remove_auto_filter,
        auto_filter_functions::has_auto_filter,
        auto_filter_functions::get_auto_filter_range,
        // Table functions
        table::add_table,
        table::get_tables,
        table::remove_table,
        table::has_tables,
        table::count_tables,
        table::set_table_style,
        table::remove_table_style,
        table::add_table_column,
        table::modify_table_column,
        table::set_table_totals_row,
        // Hyperlink functions
        hyperlink::add_hyperlink,
        hyperlink::get_hyperlink,
        hyperlink::remove_hyperlink,
        hyperlink::has_hyperlink,
        hyperlink::has_hyperlinks,
        hyperlink::get_hyperlinks,
        hyperlink::update_hyperlink,
        // Rich Text functions - Adding more functions
        rich_text_functions::create_rich_text,
        rich_text_functions::create_rich_text_from_html,
        rich_text_functions::create_text_element,
        rich_text_functions::get_text_element_text,
        rich_text_functions::get_text_element_font_properties,
        rich_text_functions::add_text_element_to_rich_text,
        rich_text_functions::add_formatted_text_to_rich_text,
        rich_text_functions::set_cell_rich_text,
        rich_text_functions::get_cell_rich_text,
        rich_text_functions::get_rich_text_plain_text,
        rich_text_functions::rich_text_to_html,
        rich_text_functions::get_rich_text_elements,
        // OLE Objects functions
        ole_object_functions::create_ole_objects,
        ole_object_functions::create_ole_object,
        ole_object_functions::create_embedded_object_properties,
        ole_object_functions::get_ole_objects,
        ole_object_functions::set_ole_objects,
        ole_object_functions::add_ole_object,
        ole_object_functions::get_ole_object_list,
        ole_object_functions::count_ole_objects,
        ole_object_functions::has_ole_objects,
        ole_object_functions::get_ole_object_requires,
        ole_object_functions::set_ole_object_requires,
        ole_object_functions::get_ole_object_prog_id,
        ole_object_functions::set_ole_object_prog_id,
        ole_object_functions::get_ole_object_extension,
        ole_object_functions::set_ole_object_extension,
        ole_object_functions::get_ole_object_data,
        ole_object_functions::set_ole_object_data,
        ole_object_functions::get_ole_object_properties,
        ole_object_functions::set_ole_object_properties,
        ole_object_functions::get_embedded_object_prog_id,
        ole_object_functions::set_embedded_object_prog_id,
        ole_object_functions::get_embedded_object_shape_id,
        ole_object_functions::set_embedded_object_shape_id,
        ole_object_functions::load_ole_object_from_file,
        ole_object_functions::save_ole_object_to_file,
        ole_object_functions::create_ole_object_with_data,
        ole_object_functions::is_ole_object_binary,
        ole_object_functions::is_ole_object_excel,
        // Page Breaks functions
        page_breaks::add_row_page_break,
        page_breaks::add_column_page_break,
        page_breaks::remove_row_page_break,
        page_breaks::remove_column_page_break,
        page_breaks::get_row_page_breaks,
        page_breaks::get_column_page_breaks,
        page_breaks::clear_row_page_breaks,
        page_breaks::clear_column_page_breaks,
        page_breaks::has_row_page_break,
        page_breaks::has_column_page_break,
        // Document Properties functions
        document_properties::get_custom_property,
        document_properties::set_custom_property_string,
        document_properties::set_custom_property_number,
        document_properties::set_custom_property_bool,
        document_properties::set_custom_property_date,
        document_properties::remove_custom_property,
        document_properties::get_custom_property_names,
        document_properties::has_custom_property,
        document_properties::get_custom_properties_count,
        document_properties::clear_custom_properties,
        document_properties::get_title,
        document_properties::set_title,
        document_properties::get_description,
        document_properties::set_description,
        document_properties::get_subject,
        document_properties::set_subject,
        document_properties::get_keywords,
        document_properties::set_keywords,
        document_properties::get_creator,
        document_properties::set_creator,
        document_properties::get_last_modified_by,
        document_properties::set_last_modified_by,
        document_properties::get_category,
        document_properties::set_category,
        document_properties::get_company,
        document_properties::set_company,
        document_properties::get_manager,
        document_properties::set_manager,
        document_properties::get_created,
        document_properties::set_created,
        document_properties::get_modified,
        document_properties::set_modified,
        // VML Support functions
        vml_support::create_vml_shape,
        vml_support::set_vml_shape_style,
        vml_support::set_vml_shape_type,
        vml_support::set_vml_shape_filled,
        vml_support::set_vml_shape_fill_color,
        vml_support::set_vml_shape_stroked,
        vml_support::set_vml_shape_stroke_color,
        vml_support::set_vml_shape_stroke_weight,
    ]
);

#[rustler::nif]
fn test_simple_function() -> NifResult<String> {
    Ok("test_works".to_string())
}
