use crate::atoms;
use crate::UmyaSpreadsheet;
use rustler::{Atom, NifResult, Error as NifError};
use std::panic::{self, AssertUnwindSafe};
use umya_spreadsheet::SheetViewValues;

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
        Ok(Err(msg)) => {
            Err(NifError::Term(Box::new((atoms::error(), msg))))
        }
        Err(_) => {
            Err(NifError::Term(Box::new((atoms::error(), "Error occurred in set_zoom_scale_normal".to_string()))))
        }
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
        Ok(Err(msg)) => {
            Err(NifError::Term(Box::new((atoms::error(), msg))))
        }
        Err(_) => {
            Err(NifError::Term(Box::new((atoms::error(), "Error occurred in set_zoom_scale_page_layout".to_string()))))
        }
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
        Ok(Err(msg)) => {
            Err(NifError::Term(Box::new((atoms::error(), msg))))
        }
        Err(_) => {
            Err(NifError::Term(Box::new((atoms::error(), "Error occurred in set_zoom_scale_page_break".to_string()))))
        }
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
                "normal" => SheetViewValues::Normal,
                "page_layout" => SheetViewValues::PageLayout,
                "page_break_preview" => SheetViewValues::PageBreakPreview,
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
        Ok(Err(msg)) => {
            Err(NifError::Term(Box::new((atoms::error(), msg))))
        }
        Err(_) => {
            Err(NifError::Term(Box::new((atoms::error(), "Error occurred in set_sheet_view".to_string()))))
        }
    }
}
