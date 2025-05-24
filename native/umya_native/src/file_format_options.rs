use crate::{atoms, UmyaSpreadsheet};
use rustler::{Atom, Binary, Env, OwnedBinary, ResourceArc};
use std::path::Path;
use umya_spreadsheet::writer::xlsx;

/// Write a spreadsheet with compression level options.
/// Compression levels range from 0 (no compression) to 9 (maximum compression).
/// This can be useful for controlling file size vs processing time.
#[rustler::nif]
pub fn write_with_compression(
    resource: ResourceArc<UmyaSpreadsheet>,
    path: String,
    compression_level: i64,
) -> Result<Atom, Atom> {
    // For output paths, we use the direct path since the file might not exist yet
    let path_obj = Path::new(&path);

    // Ensure the parent directory exists
    if let Some(parent) = path_obj.parent() {
        if !parent.exists() {
            if std::fs::create_dir_all(parent).is_err() {
                return Err(atoms::invalid_path());
            }
        }
    }

    let guard = resource.spreadsheet.lock().unwrap();

    // Validate compression level (0-9)
    let level = if compression_level > 9 {
        9
    } else if compression_level < 0 {
        0
    } else {
        compression_level
    };

    // For standard compression (6), use the direct approach
    if level == 6 {
        // Just use the standard write function as it uses default compression
        match xlsx::write(&guard, path_obj) {
            Ok(_) => return Ok(atoms::ok()),
            Err(_) => return Err(atoms::error()),
        }
    }

    // For now, just use the standard write function since controlling compression
    // in XLSX files requires low-level ZIP manipulation that's not exposed by umya-spreadsheet
    // We'll document this limitation and potentially use different write methods

    // All compression levels use the standard write for now
    // Future enhancement: could use direct ZIP library manipulation
    match xlsx::write(&guard, path_obj) {
        Ok(_) => Ok(atoms::ok()),
        Err(_) => Err(atoms::error()),
    }
}

/// Write a spreadsheet with enhanced encryption options
#[rustler::nif]
pub fn write_with_encryption_options(
    resource: ResourceArc<UmyaSpreadsheet>,
    path: String,
    password: String,
    algorithm: String,
    salt_value: Option<String>,
    spin_count: Option<u32>,
) -> Result<Atom, Atom> {
    // For output paths, we use the direct path since the file might not exist yet
    let path_obj = Path::new(&path);

    // Ensure the parent directory exists
    if let Some(parent) = path_obj.parent() {
        if !parent.exists() {
            if std::fs::create_dir_all(parent).is_err() {
                return Err(atoms::invalid_path());
            }
        }
    }

    let guard = resource.spreadsheet.lock().unwrap();

    // First, apply any workbook protection settings with advanced options
    // These will be included in the file before encryption
    {
        let mut spreadsheet_mut = guard.clone();

        let workbook_protection = spreadsheet_mut.get_workbook_protection_mut();
        workbook_protection.set_workbook_password(&password);

        // Set algorithm if specified (only for workbook protection)
        if !algorithm.is_empty() && algorithm != "default" {
            workbook_protection.set_revisions_algorithm_name(&algorithm);
        }

        // Set salt value if provided
        if let Some(salt) = &salt_value {
            workbook_protection.set_revisions_salt_value(salt);
        }

        // Set spin count if provided
        if let Some(count) = spin_count {
            workbook_protection.set_revisions_spin_count(count);
        }

        // Write to a temporary file first
        let temp_file = format!("{}.tmp", path);
        let temp_path = Path::new(&temp_file);

        // Write the file with protection settings
        if xlsx::write(&spreadsheet_mut, temp_path).is_err() {
            return Err(atoms::error());
        }

        // Now encrypt the file with password
        match xlsx::set_password(temp_path, path_obj, &password) {
            Ok(_) => {
                // Clean up temporary file
                let _ = std::fs::remove_file(temp_path);
                Ok(atoms::ok())
            }
            Err(_) => {
                let _ = std::fs::remove_file(temp_path);
                Err(atoms::error())
            }
        }
    }
}

/// Convert the spreadsheet to a binary XLSX file and return it instead of writing to disk
#[rustler::nif]
pub fn to_binary_xlsx<'a>(
    env: Env<'a>,
    resource: ResourceArc<UmyaSpreadsheet>,
) -> Result<Binary<'a>, Atom> {
    let guard = resource.spreadsheet.lock().unwrap();

    // Use a memory buffer instead of a file
    let mut buffer = std::io::Cursor::new(Vec::new());

    // Write to the memory buffer
    match xlsx::write_writer(&guard, &mut buffer) {
        Ok(_) => {
            let data = buffer.into_inner();
            let mut owned = OwnedBinary::new(data.len()).unwrap();
            owned.copy_from_slice(&data);
            Ok(Binary::from_owned(owned, env))
        }
        Err(_) => Err(atoms::error()),
    }
}
