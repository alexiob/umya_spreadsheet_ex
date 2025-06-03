use rustler::{Atom, Error as NifError, NifResult};
use std::path::Path;
use umya_spreadsheet;

use crate::atoms;

/// Set password protection on an Excel file
#[rustler::nif]
pub fn set_password(input_path: String, output_path: String, password: String) -> NifResult<Atom> {
    let input = Path::new(&input_path);
    let output = Path::new(&output_path);

    if !input.exists() {
        return Err(NifError::Term(Box::new((
            atoms::error(),
            "Input file not found".to_string(),
        ))));
    }

    match umya_spreadsheet::reader::xlsx::read(input) {
        Ok(book) => {
            match umya_spreadsheet::writer::xlsx::write_with_password(&book, output, &password) {
                Ok(_) => Ok(atoms::ok()),
                Err(_) => Err(NifError::Term(Box::new((
                    atoms::error(),
                    "Failed to write password-protected file".to_string(),
                )))),
            }
        }
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Failed to read input file".to_string(),
        )))),
    }
}
