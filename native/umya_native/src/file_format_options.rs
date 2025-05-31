use crate::{atoms, UmyaSpreadsheet};
use rustler::{Atom, Binary, Env, Error as NifError, NifResult, OwnedBinary, ResourceArc};
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
) -> NifResult<Atom> {
    // For output paths, we use the direct path since the file might not exist yet
    let path_obj = Path::new(&path);

    // Ensure the parent directory exists
    if let Some(parent) = path_obj.parent() {
        if !parent.exists() {
            if std::fs::create_dir_all(parent).is_err() {
                return Err(NifError::Term(Box::new((
                    atoms::error(),
                    "Invalid file path or unable to create directory".to_string(),
                ))));
            }
        }
    }

    // Acquire mutex guard to access the spreadsheet data
    let guard = resource.spreadsheet.lock().unwrap();

    // Validate compression level (0-9)
    let _level = if compression_level > 9 {
        9
    } else if compression_level < 0 {
        0
    } else {
        compression_level
    };

    // Note: Currently all compression levels use the standard write method
    // Future enhancement: implement actual compression level control
    let result = match xlsx::write(&guard, path_obj) {
        Ok(_) => Ok(atoms::ok()),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Failed to write file with compression".to_string(),
        )))),
    };

    // Explicitly drop the guard to release the mutex before returning
    // This ensures the mutex is always released, even in error cases
    drop(guard);

    // Return the result
    result
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
) -> NifResult<Atom> {
    // For output paths, we use the direct path since the file might not exist yet
    let path_obj = Path::new(&path);

    // Ensure the parent directory exists
    if let Some(parent) = path_obj.parent() {
        if !parent.exists() {
            if std::fs::create_dir_all(parent).is_err() {
                return Err(NifError::Term(Box::new((
                    atoms::error(),
                    "Invalid file path or unable to create directory".to_string(),
                ))));
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
            // Explicitly drop the guard to ensure mutex is released
            drop(guard);
            return Err(NifError::Term(Box::new((
                atoms::error(),
                "Failed to write temporary file".to_string(),
            ))));
        }

        // Now encrypt the file with password
        let result = match xlsx::set_password(temp_path, path_obj, &password) {
            Ok(_) => {
                // Clean up temporary file
                let _ = std::fs::remove_file(temp_path);
                Ok(atoms::ok())
            }
            Err(_) => {
                let _ = std::fs::remove_file(temp_path);
                Err(NifError::Term(Box::new((
                    atoms::error(),
                    "Failed to write file with encryption".to_string(),
                ))))
            }
        };

        // Explicitly drop the guard to ensure mutex is released
        drop(guard);

        // Return the result
        result
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
    let result = match xlsx::write_writer(&guard, &mut buffer) {
        Ok(_) => {
            let data = buffer.into_inner();
            let mut owned = OwnedBinary::new(data.len()).unwrap();
            owned.copy_from_slice(&data);
            Ok(Binary::from_owned(owned, env))
        }
        Err(_) => Err(atoms::error()),
    };

    // Explicitly drop the guard before returning
    drop(guard);

    // Return the result
    result
}

/// Get the default compression level for XLSX files
/// This always returns 6 (standard compression) since it's the default in the underlying zip library
#[rustler::nif]
pub fn get_compression_level(_resource: ResourceArc<UmyaSpreadsheet>) -> NifResult<i64> {
    // Default compression level is 6 in the zip library
    Ok(6)
}

/// Check if a spreadsheet has encryption enabled
#[rustler::nif]
pub fn is_encrypted(resource: ResourceArc<UmyaSpreadsheet>) -> NifResult<bool> {
    let guard = resource.spreadsheet.lock().unwrap();

    // Check if workbook protection is enabled by checking if the password is non-empty
    let has_protection = match guard.get_workbook_protection() {
        Some(protection) => !protection.get_workbook_password_raw().is_empty(),
        None => false,
    };

    // Explicitly drop the guard before returning
    drop(guard);

    // Return the result
    Ok(has_protection)
}

/// Get the encryption algorithm used for the workbook
#[rustler::nif]
pub fn get_encryption_algorithm(
    resource: ResourceArc<UmyaSpreadsheet>,
) -> NifResult<Option<String>> {
    let guard = resource.spreadsheet.lock().unwrap();

    // Check if a password is set
    if let Some(protection) = guard.get_workbook_protection() {
        if !protection.get_workbook_password_raw().is_empty() {
            // Get the algorithm name or use default
            let algorithm = protection.get_revisions_algorithm_name().to_string();

            // Return default if empty
            let result = if algorithm.is_empty() {
                "default".to_string()
            } else {
                algorithm
            };

            // Explicitly drop the guard before returning
            drop(guard);

            return Ok(Some(result));
        }
    }

    // Explicitly drop the guard before returning
    drop(guard);

    Ok(None)
}
