use crate::atoms;
use crate::UmyaSpreadsheet;
use rustler::{Atom, Encoder, Env, Error as NifError, NifResult, Term};
use std::panic::{self, AssertUnwindSafe};
use umya_spreadsheet::{PaneStateValues, PaneValues, SheetViewValues};

// Helper function to ensure a sheet has a sheet view with proper Excel defaults
fn ensure_sheet_view_with_defaults(sheet: &mut umya_spreadsheet::Worksheet) -> &mut umya_spreadsheet::SheetView {
    let sheet_views = sheet.get_sheet_views_mut();
    let views_list = sheet_views.get_sheet_view_list_mut();
    if views_list.is_empty() {
        views_list.push(umya_spreadsheet::SheetView::default());
    }
    
    // Get the first sheet view and ensure it has proper Excel defaults
    let sheet_view = views_list.get_mut(0).unwrap();
    
    // Check if this appears to be an unmodified default sheet view from new_file()
    // If zoom is 0 (unset) and grid lines is false, this looks like a Rust default
    let zoom_scale = *sheet_view.get_zoom_scale();
    let show_grid_lines = *sheet_view.get_show_grid_lines();
    
    if zoom_scale == 0 && !show_grid_lines {
        // This looks like an unmodified default from new_file(), set Excel defaults
        sheet_view.set_show_grid_lines(true);  // Excel default
        sheet_view.set_zoom_scale(100);        // Excel default
    } else if zoom_scale == 0 {
        // Only zoom needs fixing
        sheet_view.set_zoom_scale(100);        // Excel default
    }
    
    sheet_view
}

#[rustler::nif]
pub fn set_selection(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    active_cell: String,
    sqref: String,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) {
            // Get or create a sheet view with proper Excel defaults
            let sheet_view = ensure_sheet_view_with_defaults(sheet);

            // Create a new selection object
            let mut selection = umya_spreadsheet::Selection::default();

            // Convert string active cell to Coordinate
            let mut active_coord = umya_spreadsheet::Coordinate::default();
            active_coord.set_coordinate(&active_cell);
            selection.set_active_cell(active_coord);

            // Set sequence of references from sqref
            selection.get_sequence_of_references_mut().set_sqref(sqref);

            // Clear existing selections and add the new one
            sheet_view.get_selection_mut().clear();
            sheet_view.set_selection(selection);

            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => Err(NifError::Term(Box::new((atoms::error(), msg)))),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Error occurred in set_selection".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn set_show_grid_lines(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    value: bool,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) {
            // Get or create a sheet view with proper Excel defaults
            let sheet_view = ensure_sheet_view_with_defaults(sheet);

            // Set show grid lines value
            sheet_view.set_show_grid_lines(value);

            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => Err(NifError::Term(Box::new((atoms::error(), msg)))),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Error occurred in set_show_grid_lines".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn set_tab_selected(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    value: bool,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) {
            // Get or create a sheet view
            let sheet_views = sheet.get_sheet_views_mut();
            let views_list = sheet_views.get_sheet_view_list_mut();
            if views_list.is_empty() {
                views_list.push(umya_spreadsheet::SheetView::default());
            }
            let sheet_view = views_list.get_mut(0).unwrap();
            sheet_view.set_tab_selected(value);
            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => Err(NifError::Term(Box::new((atoms::error(), msg)))),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Error occurred in set_tab_selected".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn set_tab_color(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    color: String,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) {
            let mut tab_color = umya_spreadsheet::Color::default();
            tab_color.set_argb(color);
            sheet.set_tab_color(tab_color);
            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => Err(NifError::Term(Box::new((atoms::error(), msg)))),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Error occurred in set_tab_color".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn set_zoom_scale(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    value: u32,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) {
            // Get or create a sheet view
            let sheet_views = sheet.get_sheet_views_mut();
            let views_list = sheet_views.get_sheet_view_list_mut();
            if views_list.is_empty() {
                views_list.push(umya_spreadsheet::SheetView::default());
            }
            let sheet_view = views_list.get_mut(0).unwrap();
            sheet_view.set_zoom_scale(value);
            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => Err(NifError::Term(Box::new((atoms::error(), msg)))),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Error occurred in set_zoom_scale".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn freeze_panes(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    rows: u32,
    cols: u32,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) {
            // Get or create a sheet view
            let sheet_views = sheet.get_sheet_views_mut();
            let views_list = sheet_views.get_sheet_view_list_mut();
            if views_list.is_empty() {
                views_list.push(umya_spreadsheet::SheetView::default());
            }
            let sheet_view = views_list.get_mut(0).unwrap();

            let mut pane = umya_spreadsheet::Pane::default();
            pane.set_vertical_split(cols as f64);
            pane.set_horizontal_split(rows as f64);
            let mut top_left_coord = umya_spreadsheet::Coordinate::default();
            top_left_coord.set_coordinate(format!(
                "{}{}",
                umya_spreadsheet::helper::coordinate::string_from_column_index(&(cols + 1)),
                rows + 1
            ));
            pane.set_top_left_cell(top_left_coord);
            pane.set_active_pane(PaneValues::TopRight);
            pane.set_state(PaneStateValues::Frozen);

            sheet_view.set_pane(pane);
            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => Err(NifError::Term(Box::new((atoms::error(), msg)))),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Error occurred in freeze_panes".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn set_top_left_cell(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell: String,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) {
            // Get or create a sheet view
            let sheet_views = sheet.get_sheet_views_mut();
            let views_list = sheet_views.get_sheet_view_list_mut();
            if views_list.is_empty() {
                views_list.push(umya_spreadsheet::SheetView::default());
            }
            let sheet_view = views_list.get_mut(0).unwrap();
            sheet_view.set_top_left_cell(cell);
            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => Err(NifError::Term(Box::new((atoms::error(), msg)))),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Error occurred in set_top_left_cell".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn split_panes(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    horizontal_position: f64,
    vertical_position: f64,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) {
            // Get or create a sheet view
            let sheet_views = sheet.get_sheet_views_mut();
            let views_list = sheet_views.get_sheet_view_list_mut();
            if views_list.is_empty() {
                views_list.push(umya_spreadsheet::SheetView::default());
            }
            let sheet_view = views_list.get_mut(0).unwrap();

            let mut pane = umya_spreadsheet::Pane::default();
            pane.set_vertical_split(vertical_position); // Note: umya-spreadsheet uses x_split for vertical and y_split for horizontal
            pane.set_horizontal_split(horizontal_position);
            // Determine active pane based on split positions - this might need more logic
            // For simplicity, setting to bottom right if both splits exist
            if horizontal_position > 0.0 && vertical_position > 0.0 {
                pane.set_active_pane(PaneValues::BottomRight);
            } else if horizontal_position > 0.0 {
                pane.set_active_pane(PaneValues::BottomLeft);
            } else if vertical_position > 0.0 {
                pane.set_active_pane(PaneValues::TopRight);
            }
            // TopLeftCell is usually set automatically or based on active cell after split
            // For now, we'll leave it default or it could be set to A1 or the active cell.
            // pane.set_top_left_cell_crate("A1".to_string());
            pane.set_state(PaneStateValues::Split);
            sheet_view.set_pane(pane);
            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => Err(NifError::Term(Box::new((atoms::error(), msg)))),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Error occurred in split_panes".to_string(),
        )))),
    }
}

// Functions from sheet_view_functions_part2.rs
#[rustler::nif]
pub fn set_zoom_scale_normal(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    scale: u32,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) {
            // Get or create a sheet view
            let sheet_views = sheet.get_sheet_views_mut();
            let views_list = sheet_views.get_sheet_view_list_mut();
            if views_list.is_empty() {
                views_list.push(umya_spreadsheet::SheetView::default());
            }
            let sheet_view = views_list.get_mut(0).unwrap();

            // Set zoom scale normal
            sheet_view.set_zoom_scale_normal(scale);

            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => Err(NifError::Term(Box::new((atoms::error(), msg)))),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Error occurred in set_zoom_scale_normal".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn set_zoom_scale_page_layout(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    scale: u32,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) {
            // Get or create a sheet view
            let sheet_views = sheet.get_sheet_views_mut();
            let views_list = sheet_views.get_sheet_view_list_mut();
            if views_list.is_empty() {
                views_list.push(umya_spreadsheet::SheetView::default());
            }
            let sheet_view = views_list.get_mut(0).unwrap();

            // Set zoom scale page layout
            sheet_view.set_zoom_scale_page_layout_view(scale);

            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => Err(NifError::Term(Box::new((atoms::error(), msg)))),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Error occurred in set_zoom_scale_page_layout".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn set_zoom_scale_page_break(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    scale: u32,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) {
            // Get or create a sheet view
            let sheet_views = sheet.get_sheet_views_mut();
            let views_list = sheet_views.get_sheet_view_list_mut();
            if views_list.is_empty() {
                views_list.push(umya_spreadsheet::SheetView::default());
            }
            let sheet_view = views_list.get_mut(0).unwrap();

            // Set zoom scale sheet layout (page break preview)
            sheet_view.set_zoom_scale_sheet_layout_view(scale);

            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => Err(NifError::Term(Box::new((atoms::error(), msg)))),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Error occurred in set_zoom_scale_page_break".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn set_sheet_view(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    view_type: String,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) {
            // Get or create a sheet view
            let sheet_views = sheet.get_sheet_views_mut();
            let views_list = sheet_views.get_sheet_view_list_mut();
            if views_list.is_empty() {
                views_list.push(umya_spreadsheet::SheetView::default());
            }
            let sheet_view = views_list.get_mut(0).unwrap();

            // Set the view type based on input string
            let view_type_enum = match view_type.as_str() {
                "normal" => SheetViewValues::Normal,
                "page_layout" => SheetViewValues::PageLayout,
                "page_break_preview" => SheetViewValues::PageBreakPreview,
                _ => return Err(format!("Unsupported view type: {}", view_type)),
            };

            sheet_view.set_view(view_type_enum);

            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => Err(NifError::Term(Box::new((atoms::error(), msg)))),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Error occurred in set_sheet_view".to_string(),
        )))),
    }
}

#[rustler::nif]
pub fn get_show_grid_lines<'a>(
    env: rustler::Env<'a>,
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> rustler::Term<'a> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) {
            // Ensure the sheet view has proper Excel defaults
            let sheet_view = ensure_sheet_view_with_defaults(sheet);
            let show_grid_lines_value = *sheet_view.get_show_grid_lines();
            
            // Return the actual value from the sheet view
            (atoms::ok(), show_grid_lines_value).encode(env)
        } else {
            (atoms::error(), format!("Sheet '{}' not found", sheet_name)).encode(env)
        }
    }));

    match result {
        Ok(term) => term,
        Err(_) => (
            atoms::error(),
            "Error occurred in get_show_grid_lines"
        ).encode(env),
    }
}

#[rustler::nif]
pub fn get_zoom_scale<'a>(
    env: rustler::Env<'a>,
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> rustler::Term<'a> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name(&sheet_name) {
            // Get sheet views
            let sheet_views = sheet.get_sheets_views();
            let views_list = sheet_views.get_sheet_view_list();
            
            // If no sheet views exist, return Excel's default (100)
            if views_list.is_empty() {
                return (atoms::ok(), 100).encode(env);
            }
            
            let sheet_view = views_list.get(0).unwrap();
            let zoom_scale = sheet_view.get_zoom_scale();
            
            // If zoom scale is 0 (unset), return Excel's default
            let result_zoom = if *zoom_scale == 0 { 100 } else { *zoom_scale };
            
            (atoms::ok(), result_zoom).encode(env)
        } else {
            (atoms::error(), format!("Sheet '{}' not found", sheet_name)).encode(env)
        }
    }));

    match result {
        Ok(term) => term,
        Err(_) => (
            atoms::error(),
            "Error occurred in get_zoom_scale"
        ).encode(env),
    }
}

#[rustler::nif]
pub fn get_tab_color<'a>(
    env: rustler::Env<'a>,
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> rustler::Term<'a> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name(&sheet_name) {
            // Check if tab color is set
            match sheet.get_tab_color() {
                Some(color) => {
                    // Get the ARGB color string and remove the alpha channel (first 2 characters)
                    let argb = color.get_argb();
                    // Ensure we have the full 6-digit hex color
                    let color_hex = if argb.len() >= 8 {
                        // Take characters 2-8 to get the RGB part (skip alpha FF prefix)
                        format!("#{}", &argb[2..8])
                    } else if argb.len() >= 6 && !argb.starts_with('#') {
                        // If it's 6 chars and doesn't start with #, assume it's RGB
                        format!("#{}", &argb)
                    } else if argb.starts_with('#') {
                        // Already has #, return as-is
                        argb.to_string()
                    } else {
                        // Pad with zeros if needed
                        format!("#{:0>6}", &argb[2..])
                    };
                    (atoms::ok(), color_hex).encode(env)
                },
                None => (atoms::ok(), "").encode(env)
            }
        } else {
            (atoms::error(), format!("Sheet '{}' not found", sheet_name)).encode(env)
        }
    }));

    match result {
        Ok(term) => term,
        Err(_) => (
            atoms::error(),
            "Error occurred in get_tab_color"
        ).encode(env),
    }
}

#[rustler::nif]
pub fn get_sheet_view<'a>(
    env: rustler::Env<'a>,
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> rustler::Term<'a> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) {
            // Get or create a sheet view to ensure defaults are available
            let sheet_views = sheet.get_sheet_views_mut();
            let views_list = sheet_views.get_sheet_view_list_mut();
            if views_list.is_empty() {
                views_list.push(umya_spreadsheet::SheetView::default());
                // Set the proper Excel defaults for a new sheet view
                let sheet_view = views_list.get_mut(0).unwrap();
                sheet_view.set_show_grid_lines(true);
                sheet_view.set_zoom_scale(100);
            }
            
            let sheet_view = views_list.get(0).unwrap();
            
            // Map the view value to a readable string
            let view_type = match sheet_view.get_view() {
                SheetViewValues::Normal => "normal",
                SheetViewValues::PageBreakPreview => "page_break_preview",
                SheetViewValues::PageLayout => "page_layout",
            };
            
            (atoms::ok(), view_type).encode(env)
        } else {
            (atoms::error(), format!("Sheet '{}' not found", sheet_name)).encode(env)
        }
    }));

    match result {
        Ok(term) => term,
        Err(_) => (
            atoms::error(),
            "Error occurred in get_sheet_view"
        ).encode(env),
    }
}

#[rustler::nif]
pub fn get_selection<'a>(
    env: rustler::Env<'a>,
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> rustler::Term<'a> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) {
            // Get or create a sheet view to ensure defaults are available
            let sheet_views = sheet.get_sheet_views_mut();
            let views_list = sheet_views.get_sheet_view_list_mut();
            if views_list.is_empty() {
                views_list.push(umya_spreadsheet::SheetView::default());
                // Set the proper Excel defaults for a new sheet view
                let sheet_view = views_list.get_mut(0).unwrap();
                sheet_view.set_show_grid_lines(true);
                sheet_view.set_zoom_scale(100);
            }
            
            let sheet_view = views_list.get(0).unwrap();
            let selections = sheet_view.get_selection();
            
            // Create result map
            let mut selection_result = std::collections::HashMap::new();
            
            // Check if we have any selections
            if !selections.is_empty() {
                let selection = &selections[0]; // Get first selection
                
                // Get active cell - handle Option
                if let Some(active_cell_coord) = selection.get_active_cell() {
                    let active_cell = active_cell_coord.get_coordinate();
                    selection_result.insert("active_cell".to_string(), active_cell);
                } else {
                    selection_result.insert("active_cell".to_string(), "A1".to_string());
                }
                
                // Get sqref (selection range)
                let sqref = selection.get_sequence_of_references().get_sqref().clone();
                selection_result.insert("sqref".to_string(), sqref);
            } else {
                // Default values if no selection
                selection_result.insert("active_cell".to_string(), "A1".to_string());
                selection_result.insert("sqref".to_string(), "A1".to_string());
            }
            
            (atoms::ok(), selection_result).encode(env)
        } else {
            (atoms::error(), format!("Sheet '{}' not found", sheet_name)).encode(env)
        }
    }));

    match result {
        Ok(term) => term,
        Err(_) => (
            atoms::error(),
            "Error occurred in get_selection"
        ).encode(env),
    }
}
