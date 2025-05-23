use crate::{atoms, helpers::alignment_helper, UmyaSpreadsheet};
use rustler::{Atom, ResourceArc};

#[rustler::nif]
fn set_cell_alignment(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    horizontal: String,
    vertical: String,
) -> Result<Atom, Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let cell = sheet.get_cell_mut(&*cell_address);

            // Set horizontal alignment if provided
            if !horizontal.is_empty() {
                let horizontal_alignment =
                    alignment_helper::parse_horizontal_alignment(horizontal.as_str());
                cell.get_style_mut()
                    .get_alignment_mut()
                    .set_horizontal(horizontal_alignment);
            }

            // Set vertical alignment if provided
            if !vertical.is_empty() {
                let vertical_alignment =
                    alignment_helper::parse_vertical_alignment(vertical.as_str());
                cell.get_style_mut()
                    .get_alignment_mut()
                    .set_vertical(vertical_alignment);
            }

            Ok(atoms::ok())
        }
        None => Err(atoms::not_found()),
    }
}
