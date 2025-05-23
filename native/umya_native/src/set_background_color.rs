use rustler::{Atom, ResourceArc};
use crate::{UmyaSpreadsheet, atoms, helpers};

#[rustler::nif]
fn set_background_color(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    color: String,
) -> Result<Atom, Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            // Use the color helper to create a Color object
            let color_obj = match helpers::color_helper::create_color_object(&color) {
                Ok(color) => color,
                Err(e) => return Err(e),
            };

            // Apply the background color using the style helper
            helpers::style_helper::apply_cell_style(
                sheet,
                &*cell_address,
                Some(color_obj),
                None,
                None,
                None,
            );

            Ok(atoms::ok())
        }
        None => Err(atoms::not_found()),
    }
}
