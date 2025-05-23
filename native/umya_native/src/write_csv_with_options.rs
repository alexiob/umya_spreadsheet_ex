use crate::{atoms, helpers::format_helper, UmyaSpreadsheet};
use rustler::{Atom, ResourceArc};
use std::path::Path;

#[rustler::nif]
fn write_csv(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    path: String,
) -> Result<Atom, Atom> {
    // Call the implementation directly with default values
    write_csv_with_options_impl(
        resource,
        sheet_name,
        path,
        "UTF8".to_string(), // Default UTF-8 encoding
        ",".to_string(),    // Default comma delimiter
        false,              // Default no trimming
        "".to_string(),     // Default no wrap character
    )
}

#[rustler::nif]
fn write_csv_with_options(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    path: String,
    encoding: String,
    delimiter: String,
    do_trim: bool,
    wrap_with_char: String,
) -> Result<Atom, Atom> {
    write_csv_with_options_impl(
        resource,
        sheet_name,
        path,
        encoding,
        delimiter,
        do_trim,
        wrap_with_char,
    )
}

// Implementation function to avoid code duplication
fn write_csv_with_options_impl(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    path: String,
    encoding: String,
    delimiter: String,
    do_trim: bool,
    wrap_with_char: String,
) -> Result<Atom, Atom> {
    let guard = resource.spreadsheet.lock().unwrap();

    // For output paths, we use the direct path since the file might not exist yet
    // We want to ensure the directory exists though
    let path_obj = Path::new(&path);

    // Ensure the parent directory exists
    if let Some(parent) = path_obj.parent() {
        if !parent.exists() {
            if let Err(_) = std::fs::create_dir_all(parent) {
                return Err(atoms::invalid_path());
            }
        }
    }

    // Find the specified sheet
    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            // Create a new spreadsheet with only this sheet for export
            let mut export_spreadsheet = umya_spreadsheet::new_file_empty_worksheet();

            // Create a clone of the sheet and add it to the export spreadsheet
            let mut sheet_clone = sheet.clone();
            sheet_clone.set_name("Sheet1"); // Use a standard name for CSV export

            // Add the sheet to the export spreadsheet
            match export_spreadsheet.add_sheet(sheet_clone) {
                Ok(_) => {
                    // Create CSV options using helper
                    let option = format_helper::create_csv_writer_options(
                        &encoding,
                        do_trim,
                        &wrap_with_char,
                    );

                    // Custom implementation to handle delimiter since the underlying lib doesn't support it directly
                    if !delimiter.is_empty() {
                        // Get the active sheet
                        let worksheet = export_spreadsheet.get_active_sheet();

                        // Get max column and row
                        let (max_column, max_row) = worksheet.get_highest_column_and_row();

                        let mut data = String::new();
                        for row in 0u32..max_row {
                            let mut row_vec: Vec<String> = Vec::new();
                            for column in 0u32..max_column {
                                // Get value
                                let mut value = match worksheet.get_cell((column + 1, row + 1)) {
                                    Some(cell) => cell.get_cell_value().get_value().into(),
                                    None => String::new(),
                                };

                                // Apply trimming if needed
                                if *option.get_do_trim() {
                                    value = value.trim().to_string();
                                }

                                // Apply wrapping if needed
                                if option.get_wrap_with_char() != "" {
                                    value = format!(
                                        "{}{}{}",
                                        option.get_wrap_with_char(),
                                        value,
                                        option.get_wrap_with_char()
                                    );
                                }

                                row_vec.push(value);
                            }

                            // Use the provided delimiter instead of a hardcoded comma
                            data.push_str(&row_vec.join(&delimiter));
                            data.push_str("\r\n");
                        }

                        // Write the data to a file
                        std::fs::write(path_obj, data).map_err(|_| atoms::error())?;

                        Ok(atoms::ok())
                    } else {
                        // Fall back to the library's implementation for the default comma delimiter
                        match umya_spreadsheet::writer::csv::write(
                            &export_spreadsheet,
                            path_obj,
                            Some(&option),
                        ) {
                            Ok(_) => Ok(atoms::ok()),
                            Err(_) => Err(atoms::error()),
                        }
                    }
                }
                Err(_) => Err(atoms::error()),
            }
        }
        None => Err(atoms::not_found()),
    }
}
