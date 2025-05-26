use crate::atoms;
use crate::UmyaSpreadsheet;
use rustler::{Atom, Error as NifError, NifResult, ResourceArc};
use std::collections::HashMap;
use umya_spreadsheet::Hyperlink;

/// Add a hyperlink to a cell
///
/// # Arguments
/// * `resource` - The spreadsheet resource
/// * `sheet_name` - Name of the worksheet
/// * `cell` - Cell reference (e.g., "A1")
/// * `url` - The URL or file path for the hyperlink
/// * `tooltip` - Optional tooltip text (can be empty string)
/// * `is_internal` - Whether this is an internal worksheet reference
///
/// # Returns
/// * `:ok` on success
/// * `{:error, reason}` on failure
#[rustler::nif]
pub fn add_hyperlink(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell: String,
    url: String,
    tooltip: Option<String>,
    is_internal: bool,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(worksheet) => {
            let cell_obj = worksheet.get_cell_mut(&*cell);

            // Create hyperlink
            let mut hyperlink = Hyperlink::default();
            hyperlink.set_url(url);
            if let Some(tooltip_text) = tooltip {
                if !tooltip_text.is_empty() {
                    hyperlink.set_tooltip(tooltip_text);
                }
            }
            if is_internal {
                hyperlink.set_location(true);
            }

            // Add hyperlink to cell
            cell_obj.set_hyperlink(hyperlink);

            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Get hyperlink information from a cell
///
/// # Arguments
/// * `resource` - The spreadsheet resource
/// * `sheet_name` - Name of the worksheet
/// * `cell` - Cell reference (e.g., "A1")
///
/// # Returns
/// * `{:ok, hyperlink_info}` with map containing url, tooltip, location
/// * `{:error, reason}` on failure or if no hyperlink exists
#[rustler::nif]
pub fn get_hyperlink(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell: String,
) -> NifResult<(Atom, HashMap<String, String>)> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(worksheet) => {
            let cell_obj = worksheet.get_cell(&*cell);

            match cell_obj {
                Some(cell) => match cell.get_hyperlink() {
                    Some(hyperlink) => {
                        // Check if hyperlink has meaningful content (non-empty URL)
                        if hyperlink.get_url().is_empty() {
                            return Err(NifError::Term(Box::new((
                                atoms::error(),
                                "No hyperlink found in cell".to_string(),
                            ))));
                        }

                        let mut hyperlink_map = HashMap::new();
                        hyperlink_map.insert("url".to_string(), hyperlink.get_url().to_string());
                        hyperlink_map
                            .insert("tooltip".to_string(), hyperlink.get_tooltip().to_string());
                        // For location field, return the cell address where the hyperlink is located
                        let cell_address =
                            umya_spreadsheet::helper::coordinate::coordinate_from_index(
                                cell.get_coordinate().get_col_num(),
                                cell.get_coordinate().get_row_num(),
                            );
                        hyperlink_map.insert("location".to_string(), cell_address);

                        Ok((atoms::ok(), hyperlink_map))
                    }
                    None => Err(NifError::Term(Box::new((
                        atoms::error(),
                        "No hyperlink found in cell".to_string(),
                    )))),
                },
                None => Err(NifError::Term(Box::new((
                    atoms::error(),
                    "Cell not found".to_string(),
                )))),
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Remove hyperlink from a cell
///
/// # Arguments
/// * `resource` - The spreadsheet resource
/// * `sheet_name` - Name of the worksheet
/// * `cell` - Cell reference (e.g., "A1")
///
/// # Returns
/// * `:ok` on success
/// * `{:error, reason}` on failure
#[rustler::nif]
pub fn remove_hyperlink(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell: String,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(worksheet) => {
            let cell_obj = worksheet.get_cell_mut(&*cell);

            // Remove hyperlink by setting it to an empty hyperlink
            let empty_hyperlink = Hyperlink::default();
            cell_obj.set_hyperlink(empty_hyperlink);

            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Check if a cell has a hyperlink
///
/// # Arguments
/// * `resource` - The spreadsheet resource
/// * `sheet_name` - Name of the worksheet
/// * `cell` - Cell reference (e.g., "A1")
///
/// # Returns
/// * `{:ok, true/false}` indicating if hyperlink exists
/// * `{:error, reason}` on failure
#[rustler::nif]
pub fn has_hyperlink(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell: String,
) -> NifResult<(Atom, bool)> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(worksheet) => {
            let cell_obj = worksheet.get_cell(&*cell);
            let has_link = match cell_obj {
                Some(cell) => match cell.get_hyperlink() {
                    Some(hyperlink) => !hyperlink.get_url().is_empty(),
                    None => false,
                },
                None => false,
            };
            Ok((atoms::ok(), has_link))
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Check if a worksheet has any hyperlinks
///
/// # Arguments
/// * `resource` - The spreadsheet resource
/// * `sheet_name` - Name of the worksheet
///
/// # Returns
/// * `{:ok, true/false}` indicating if any hyperlinks exist
/// * `{:error, reason}` on failure
#[rustler::nif]
pub fn has_hyperlinks(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> NifResult<(Atom, bool)> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(worksheet) => {
            // Use public cell collection method to check for hyperlinks
            let has_links =
                worksheet
                    .get_cell_collection()
                    .iter()
                    .any(|cell| match cell.get_hyperlink() {
                        Some(hyperlink) => !hyperlink.get_url().is_empty(),
                        None => false,
                    });
            Ok((atoms::ok(), has_links))
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Get all hyperlinks from a worksheet
///
/// # Arguments
/// * `resource` - The spreadsheet resource
/// * `sheet_name` - Name of the worksheet
///
/// # Returns
/// * `{:ok, hyperlinks_list}` with list of maps containing cell, url, tooltip, location
/// * `{:error, reason}` on failure
#[rustler::nif]
pub fn get_hyperlinks(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> NifResult<(Atom, Vec<HashMap<String, String>>)> {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(worksheet) => {
            // Manually implement hyperlink collection to hashmap since the method is private
            let mut hyperlinks_list = Vec::new();

            // Iterate through all cells to find hyperlinks
            for cell in worksheet.get_cell_collection() {
                if let Some(hyperlink) = cell.get_hyperlink() {
                    // Only include hyperlinks with non-empty URLs
                    if !hyperlink.get_url().is_empty() {
                        let mut hyperlink_info = HashMap::new();

                        // Convert cell coordinate to string - use the same format as get_hyperlink
                        let cell_address =
                            umya_spreadsheet::helper::coordinate::coordinate_from_index(
                                cell.get_coordinate().get_col_num(),
                                cell.get_coordinate().get_row_num(),
                            );

                        hyperlink_info.insert("cell".to_string(), cell_address.clone());
                        hyperlink_info.insert("url".to_string(), hyperlink.get_url().to_string());
                        hyperlink_info
                            .insert("tooltip".to_string(), hyperlink.get_tooltip().to_string());
                        hyperlink_info.insert("location".to_string(), cell_address);

                        hyperlinks_list.push(hyperlink_info);
                    }
                }
            }

            Ok((atoms::ok(), hyperlinks_list))
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Update an existing hyperlink in a cell
///
/// # Arguments
/// * `resource` - The spreadsheet resource
/// * `sheet_name` - Name of the worksheet
/// * `cell` - Cell reference (e.g., "A1")
/// * `url` - The new URL or file path for the hyperlink
/// * `tooltip` - New tooltip text (can be empty string)
/// * `is_internal` - Whether this is an internal worksheet reference
///
/// # Returns
/// * `:ok` on success
/// * `{:error, reason}` on failure
#[rustler::nif]
pub fn update_hyperlink(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell: String,
    url: String,
    tooltip: Option<String>,
    is_internal: bool,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(worksheet) => {
            let cell_obj = worksheet.get_cell_mut(&*cell);

            // Check if hyperlink exists
            if cell_obj.get_hyperlink().is_none() {
                return Err(NifError::Term(Box::new((
                    atoms::error(),
                    "No hyperlink found in cell to update".to_string(),
                ))));
            }

            // Update the hyperlink by creating a new one
            let mut hyperlink = Hyperlink::default();
            hyperlink.set_url(url);
            if let Some(tooltip_text) = tooltip {
                if !tooltip_text.is_empty() {
                    hyperlink.set_tooltip(tooltip_text);
                }
            }
            if is_internal {
                hyperlink.set_location(true);
            }

            cell_obj.set_hyperlink(hyperlink);

            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}
