use rustler::{Atom, Error as NifError, NifResult, ResourceArc};
use std::path::Path;
use umya_spreadsheet::structs::drawing::spreadsheet::MarkerType;
use umya_spreadsheet::structs::Image;

use crate::atoms;
use crate::UmyaSpreadsheet;

/// Helper function to ensure a worksheet is deserialized
fn ensure_worksheet_deserialized(
    guard: &mut std::sync::MutexGuard<umya_spreadsheet::Spreadsheet>,
    sheet_index: &usize,
) {
    // Call read_sheet to deserialize the worksheet
    // This is safe to call even if already deserialized
    guard.read_sheet(*sheet_index);
}

/// Add an image to a sheet
#[rustler::nif]
pub fn add_image(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    image_path: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let path = Path::new(&image_path);
            if !path.exists() {
                return Err(NifError::Term(Box::new((
                    atoms::error(),
                    "Image not found".to_string(),
                ))));
            }

            // Create a new image
            let mut image = Image::default();

            // Create marker and set coordinate
            let mut marker = MarkerType::default();
            marker.set_coordinate(&cell_address);

            // Add the image to the Image struct using umya-spreadsheet API
            image.new_image(&image_path, marker);

            // Add the image to the worksheet
            sheet.add_image(image);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Download an image from a sheet
#[rustler::nif]
pub fn download_image(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    output_path: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Check if there's an image at this location using umya-spreadsheet API
            match sheet.get_image(cell_address.as_str()) {
                Some(image) => {
                    // Check if the image actually has data
                    if !image.has_image() {
                        return Err(NifError::Term(Box::new((
                            atoms::error(),
                            "Image not found".to_string(),
                        ))));
                    }

                    // Create the directory if it doesn't exist
                    if let Some(parent) = Path::new(&output_path).parent() {
                        if let Err(_) = std::fs::create_dir_all(parent) {
                            return Err(NifError::Term(Box::new((
                                atoms::error(),
                                "Failed to create output directory".to_string(),
                            ))));
                        }
                    }

                    // Use umya-spreadsheet's download_image method
                    image.download_image(&output_path);
                    Ok(atoms::ok())
                }
                None => Err(NifError::Term(Box::new((
                    atoms::error(),
                    "Image not found".to_string(),
                )))),
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Change an image in a cell
#[rustler::nif]
pub fn change_image(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    new_image_path: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Check if the new image file exists
            let path = Path::new(&new_image_path);
            if !path.exists() {
                return Err(NifError::Term(Box::new((
                    atoms::error(),
                    "Image not found".to_string(),
                ))));
            }

            // Check if there's an image at this location
            match sheet.get_image_mut(cell_address.as_str()) {
                Some(image) => {
                    // Check if the image actually has data
                    if !image.has_image() {
                        return Err(NifError::Term(Box::new((
                            atoms::error(),
                            "Image not found".to_string(),
                        ))));
                    }

                    // Use umya-spreadsheet's change_image method
                    image.change_image(&new_image_path);
                    Ok(atoms::ok())
                }
                None => Err(NifError::Term(Box::new((
                    atoms::error(),
                    "Image not found".to_string(),
                )))),
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Get dimensions (width, height) of an image in a cell
#[rustler::nif]
pub fn get_image_dimensions(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<(u32, u32)> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    // Find sheet index by name
    let sheet_index = match guard
        .get_sheet_collection_no_check()
        .iter()
        .position(|sheet| sheet.get_name() == sheet_name)
    {
        Some(index) => index,
        None => {
            return Err(NifError::Term(Box::new((
                atoms::error(),
                "Sheet not found".to_string(),
            ))))
        }
    };

    // Ensure the worksheet is deserialized before accessing it
    ensure_worksheet_deserialized(&mut guard, &sheet_index);

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            match sheet.get_image(cell_address.as_str()) {
                Some(image) => {
                    // Extract width and height from the image
                    let mut width = 0;
                    let mut height = 0;

                    // Width and height are stored in the extent of the anchor
                    if let Some(anchor) = image.get_one_cell_anchor() {
                        // Convert from EMU (English Metric Units) to pixels
                        // 1 pixel = 9525 EMU
                        width = (anchor.get_extent().get_cx() / 9525) as u32;
                        height = (anchor.get_extent().get_cy() / 9525) as u32;
                    } else if let Some(_anchor) = image.get_two_cell_anchor() {
                        // For two cell anchors, calculate from the from/to markers
                        let from_col = *image.get_from_marker_type().get_col();
                        let from_row = *image.get_from_marker_type().get_row();

                        if let Some(to_marker) = image.get_to_marker_type() {
                            let to_col = *to_marker.get_col();
                            let to_row = *to_marker.get_row();

                            // Simple difference estimation (this is approximate)
                            width = (to_col - from_col) * 64; // Approximate cell width
                            height = (to_row - from_row) * 20; // Approximate cell height
                        }
                    }

                    Ok((width, height))
                }
                None => Err(NifError::Term(Box::new((
                    atoms::error(),
                    "Image not found".to_string(),
                )))),
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// List all images in a sheet with their coordinates
#[rustler::nif]
pub fn list_images(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> NifResult<Vec<(String, String)>> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    // Find sheet index by name
    let sheet_index = match guard
        .get_sheet_collection_no_check()
        .iter()
        .position(|sheet| sheet.get_name() == sheet_name)
    {
        Some(index) => index,
        None => {
            return Err(NifError::Term(Box::new((
                atoms::error(),
                "Sheet not found".to_string(),
            ))))
        }
    };

    // Ensure the worksheet is deserialized before accessing it
    ensure_worksheet_deserialized(&mut guard, &sheet_index);

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            let image_collection = sheet.get_image_collection();
            let mut result = Vec::new();

            for image in image_collection {
                // Get coordinate (cell reference like "A1")
                let coordinate = image.get_coordinate();

                // Get image name
                let name = image.get_image_name().to_string();

                result.push((coordinate, name));
            }

            Ok(result)
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Get comprehensive information about an image
#[rustler::nif]
pub fn get_image_info(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<(String, String, u32, u32)> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    // Find sheet index by name
    let sheet_index = match guard
        .get_sheet_collection_no_check()
        .iter()
        .position(|sheet| sheet.get_name() == sheet_name)
    {
        Some(index) => index,
        None => {
            return Err(NifError::Term(Box::new((
                atoms::error(),
                "Sheet not found".to_string(),
            ))))
        }
    };

    // Ensure the worksheet is deserialized before accessing it
    ensure_worksheet_deserialized(&mut guard, &sheet_index);

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            match sheet.get_image(cell_address.as_str()) {
                Some(image) => {
                    // Get basic info
                    let name = image.get_image_name().to_string();
                    let position = image.get_coordinate();

                    // Get dimensions
                    let mut width = 0;
                    let mut height = 0;

                    // Width and height are stored in the extent of the anchor
                    if let Some(anchor) = image.get_one_cell_anchor() {
                        // Convert from EMU (English Metric Units) to pixels
                        // 1 pixel = 9525 EMU
                        width = (anchor.get_extent().get_cx() / 9525) as u32;
                        height = (anchor.get_extent().get_cy() / 9525) as u32;
                    } else if let Some(_anchor) = image.get_two_cell_anchor() {
                        // For two cell anchors, calculate from the from/to markers
                        let from_col = *image.get_from_marker_type().get_col();
                        let from_row = *image.get_from_marker_type().get_row();

                        if let Some(to_marker) = image.get_to_marker_type() {
                            let to_col = *to_marker.get_col();
                            let to_row = *to_marker.get_row();

                            // Simple difference estimation (this is approximate)
                            width = (to_col - from_col) * 64; // Approximate cell width
                            height = (to_row - from_row) * 20; // Approximate cell height
                        }
                    }

                    Ok((name, position, width, height))
                }
                None => Err(NifError::Term(Box::new((
                    atoms::error(),
                    "Image not found".to_string(),
                )))),
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}
