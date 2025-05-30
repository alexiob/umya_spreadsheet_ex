use rustler::{Atom, Error as NifError, NifResult, ResourceArc};

use crate::atoms;
use crate::UmyaSpreadsheet;

/// Set wrap text for a cell
#[rustler::nif]
pub fn set_wrap_text(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    wrap: bool,
) -> NifResult<Atom> {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let cell = sheet.get_cell_mut(cell_address.as_str());
            cell.get_style_mut().get_alignment_mut().set_wrap_text(wrap);
            Ok(atoms::ok())
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}
