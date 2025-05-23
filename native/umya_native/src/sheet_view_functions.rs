use crate::atoms;
use crate::helpers::error_helper::handle_error;
use crate::UmyaSpreadsheet;
use rustler::{Atom, NifResult};
use std::panic::{self, AssertUnwindSafe};
use umya_spreadsheet::{PaneStateValues, PaneValues};

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
            // Get or create a sheet view
            let sheet_views = sheet.get_sheet_views_mut();
            let views_list = sheet_views.get_sheet_view_list_mut();
            if views_list.is_empty() {
                views_list.push(umya_spreadsheet::SheetView::default());
            }
            let sheet_view = views_list.get_mut(0).unwrap();

            // Set show grid lines value
            sheet_view.set_show_grid_lines(value);

            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => handle_error(&msg),
        Err(_) => handle_error("Panic occurred in set_show_gridlines"),
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

            // Set tab selected value
            sheet_view.set_tab_selected(value);

            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => handle_error(&msg),
        Err(_) => handle_error("Panic occurred in set_tab_selected"),
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
            // Create and set the tab color
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
        Ok(Err(msg)) => handle_error(&msg),
        Err(_) => handle_error("Panic occurred in set_tab_color"),
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

            // Set zoom scale value
            sheet_view.set_zoom_scale(value);

            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => handle_error(&msg),
        Err(_) => handle_error("Panic occurred in set_zoom_scale"),
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

            // Create and configure pane
            let mut pane = umya_spreadsheet::Pane::default();
            // Set split coordinates
            pane.set_vertical_split(rows as f64);
            pane.set_horizontal_split(cols as f64);
            pane.set_active_pane(PaneValues::BottomRight);
            pane.set_state(PaneStateValues::Frozen);
            sheet_view.set_pane(pane);

            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => handle_error(&msg),
        Err(_) => handle_error("Panic occurred in freeze_panes"),
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

            // Set top left cell value
            sheet_view.set_top_left_cell(cell);

            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => handle_error(&msg),
        Err(_) => handle_error("Panic occurred in set_top_left_cell"),
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
            let view_type = match view_type.as_str() {
                "normal" => umya_spreadsheet::ViewValues::Normal,
                "page_layout" => umya_spreadsheet::ViewValues::PageLayout,
                "page_break_preview" => umya_spreadsheet::ViewValues::PageBreakPreview,
                _ => return Err(format!("Unsupported view type: {}", view_type)),
            };

            sheet_view.set_view(view_type);

            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => handle_error(&msg),
        Err(_) => handle_error("Panic occurred in set_sheet_view"),
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

            // Create and configure pane
            let mut pane = umya_spreadsheet::Pane::default();
            // Set split coordinates
            pane.set_horizontal_split(horizontal_position);
            pane.set_vertical_split(vertical_position);
            pane.set_active_pane(PaneValues::BottomRight);
            pane.set_state(PaneStateValues::Split);
            sheet_view.set_pane(pane);

            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => handle_error(&msg),
        Err(_) => handle_error("Panic occurred in split_panes"),
    }
}
