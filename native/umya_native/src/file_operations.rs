use rustler::{Atom, Error as NifError, NifResult, ResourceArc};
use std::path::Path;
use umya_spreadsheet;

use crate::atoms;
use crate::helpers;
use crate::UmyaSpreadsheet;

/// Create a new spreadsheet file with default sheet
#[rustler::nif]
pub fn new_file() -> NifResult<ResourceArc<UmyaSpreadsheet>> {
    let spreadsheet = umya_spreadsheet::new_file();
    let resource = ResourceArc::new(UmyaSpreadsheet {
        spreadsheet: std::sync::Mutex::new(spreadsheet),
    });
    Ok(resource)
}

/// Create a new spreadsheet file without any sheets
#[rustler::nif]
pub fn new_file_empty_worksheet() -> NifResult<ResourceArc<UmyaSpreadsheet>> {
    let spreadsheet = umya_spreadsheet::new_file_empty_worksheet();
    let resource = ResourceArc::new(UmyaSpreadsheet {
        spreadsheet: std::sync::Mutex::new(spreadsheet),
    });
    Ok(resource)
}

/// Read a spreadsheet file with full loading
#[rustler::nif]
pub fn read_file(path: String) -> NifResult<ResourceArc<UmyaSpreadsheet>> {
    // Use path helper to find a valid file path
    let valid_path = match helpers::path_helper::find_valid_file_path(&path) {
        Some(p) => p,
        None => {
            return Err(NifError::Term(Box::new((
                atoms::error(),
                "File not found".to_string(),
            ))))
        }
    };

    let path_obj = Path::new(&valid_path);

    // Improved error handling with specific error messages for corrupted files
    match umya_spreadsheet::reader::xlsx::read(path_obj) {
        Ok(spreadsheet) => {
            let resource = ResourceArc::new(UmyaSpreadsheet {
                spreadsheet: std::sync::Mutex::new(spreadsheet),
            });
            Ok(resource)
        }
        Err(e) => {
            // Provide specific error messages based on the error type
            let error_msg = match e.to_string().as_str() {
                s if s.contains("zip") => "corrupted_file",
                s if s.contains("xml") => "invalid_format",
                s if s.contains("permission") => "access_denied",
                _ => "read_error",
            };
            Err(NifError::Term(Box::new((
                atoms::error(),
                error_msg.to_string(),
            ))))
        }
    }
}

/// Read a spreadsheet file with lazy loading
#[rustler::nif]
pub fn lazy_read_file(path: String) -> NifResult<ResourceArc<UmyaSpreadsheet>> {
    // Use path helper to find a valid file path
    let valid_path = match helpers::path_helper::find_valid_file_path(&path) {
        Some(p) => p,
        None => {
            return Err(NifError::Term(Box::new((
                atoms::error(),
                "File not found".to_string(),
            ))))
        }
    };

    let path_obj = Path::new(&valid_path);

    // Handle both .xlsx and .xlsm files with lazy loading and improved error handling
    match umya_spreadsheet::reader::xlsx::lazy_read(path_obj) {
        Ok(spreadsheet) => {
            let resource = ResourceArc::new(UmyaSpreadsheet {
                spreadsheet: std::sync::Mutex::new(spreadsheet),
            });
            Ok(resource)
        }
        Err(e) => {
            // Provide specific error messages based on the error type
            let error_msg = match e.to_string().as_str() {
                s if s.contains("zip") => "corrupted_file",
                s if s.contains("xml") => "invalid_format",
                s if s.contains("permission") => "access_denied",
                _ => "read_error",
            };
            Err(NifError::Term(Box::new((
                atoms::error(),
                error_msg.to_string(),
            ))))
        }
    }
}

/// Write a spreadsheet to a file
#[rustler::nif]
pub fn write_file(resource: ResourceArc<UmyaSpreadsheet>, path: String) -> NifResult<Atom> {
    // For output paths, we use the direct path since the file might not exist yet
    // We want to ensure the directory exists though
    let path_obj = Path::new(&path);

    // Ensure the parent directory exists
    if let Some(parent) = path_obj.parent() {
        if !parent.exists() {
            if let Err(_) = std::fs::create_dir_all(parent) {
                return Err(NifError::Term(Box::new((
                    atoms::error(),
                    "failed_to_create_directory".to_string(),
                ))));
            }
        }
    }

    let guard = resource.spreadsheet.lock().unwrap();

    match umya_spreadsheet::writer::xlsx::write(&guard, path_obj) {
        Ok(_) => Ok(atoms::ok()),
        Err(e) => {
            let error_msg = match e.to_string().as_str() {
                s if s.contains("permission") => "access_denied",
                s if s.contains("space") => "insufficient_disk_space",
                _ => "write_error",
            };
            Err(NifError::Term(Box::new((
                atoms::error(),
                error_msg.to_string(),
            ))))
        }
    }
}

/// Write a spreadsheet to a file, using light mode
#[rustler::nif]
pub fn write_file_light(
    resource: ResourceArc<UmyaSpreadsheet>,
    path: String,
) -> Result<Atom, (Atom, String)> {
    // For output paths, we use the direct path since the file might not exist yet
    // We want to ensure the directory exists though
    let path_obj = Path::new(&path);

    // Ensure the parent directory exists
    if let Some(parent) = path_obj.parent() {
        if !parent.exists() {
            if let Err(_) = std::fs::create_dir_all(parent) {
                return Err((atoms::error(), "failed_to_create_directory".to_string()));
            }
        }
    }

    let guard = resource.spreadsheet.lock().unwrap();

    match umya_spreadsheet::writer::xlsx::write_light(&guard, path_obj) {
        Ok(_) => Ok(atoms::ok()),
        Err(e) => {
            let error_msg = match e.to_string().as_str() {
                s if s.contains("permission") => "access_denied",
                s if s.contains("space") => "insufficient_disk_space",
                _ => "write_error",
            };
            Err((atoms::error(), error_msg.to_string()))
        }
    }
}

/// Write a spreadsheet to a file with password protection
#[rustler::nif]
pub fn write_file_with_password(
    resource: ResourceArc<UmyaSpreadsheet>,
    path: String,
    password: String,
) -> NifResult<Atom> {
    // For output paths, we use the direct path since the file might not exist yet
    // We want to ensure the directory exists though
    let path_obj = Path::new(&path);

    // Ensure the parent directory exists
    if let Some(parent) = path_obj.parent() {
        if !parent.exists() {
            if let Err(_) = std::fs::create_dir_all(parent) {
                return Err(NifError::Term(Box::new((
                    atoms::error(),
                    "Invalid file path or unable to create directory".to_string(),
                ))));
            }
        }
    }

    let guard = resource.spreadsheet.lock().unwrap();

    let result =
        match umya_spreadsheet::writer::xlsx::write_with_password(&guard, path_obj, &password) {
            Ok(_) => Ok(atoms::ok()),
            Err(_) => Err(NifError::Term(Box::new((
                atoms::error(),
                "Failed to write file with password".to_string(),
            )))),
        };

    // Explicitly drop the guard to ensure mutex is released before returning
    drop(guard);

    // Return the result
    result
}

/// Write a spreadsheet to a file with password protection, using light mode
#[rustler::nif]
pub fn write_file_with_password_light(
    resource: ResourceArc<UmyaSpreadsheet>,
    path: String,
    password: String,
) -> NifResult<Atom> {
    // For output paths, we use the direct path since the file might not exist yet
    // We want to ensure the directory exists though
    let path_obj = Path::new(&path);

    // Ensure the parent directory exists
    if let Some(parent) = path_obj.parent() {
        if !parent.exists() {
            if let Err(_) = std::fs::create_dir_all(parent) {
                return Err(NifError::Term(Box::new((
                    atoms::error(),
                    "Invalid file path or unable to create directory".to_string(),
                ))));
            }
        }
    }

    let guard = resource.spreadsheet.lock().unwrap();

    let result = match umya_spreadsheet::writer::xlsx::write_with_password_light(
        &guard, path_obj, &password,
    ) {
        Ok(_) => Ok(atoms::ok()),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Failed to write file with password".to_string(),
        )))),
    };

    // Explicitly drop the guard to ensure mutex is released before returning
    drop(guard);

    // Return the result
    result
}
