use crate::atoms;
use crate::helpers::error_helper::handle_error;
use crate::UmyaSpreadsheet;
use rustler::{Atom, NifResult};
use std::panic::{self, AssertUnwindSafe};
use umya_spreadsheet::Comment;

#[rustler::nif]
pub fn add_comment(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    text: String,
    author: String,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) {
            // Create a new comment
            let mut comment = Comment::default();

            // Set up the comment with cell address, text and author
            comment.new_comment(cell_address);
            comment.set_text_string(text);
            comment.set_author(author);

            // Add the comment to the worksheet
            sheet.add_comments(comment);

            Ok(atoms::ok())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => handle_error(&msg),
        Err(_) => handle_error("Panic occurred in add_comment"),
    }
}

#[rustler::nif]
pub fn get_comment(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<(Atom, String, String)> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) {
            // Get comments as hashmap for easy lookup
            let comments = sheet.get_comments_to_hashmap();

            // Find comment for the specified cell address
            if let Some(comment) = comments.get(&cell_address) {
                // Return comment text and author
                Ok((
                    atoms::ok(),
                    comment.get_text().get_text().to_string(),
                    comment.get_author().to_string(),
                ))
            } else {
                Err(format!("No comment found at cell '{}'", cell_address))
            }
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(result)) => Ok(result),
        Ok(Err(msg)) => {
            let error_atom = handle_error(&msg)?;
            // Convert to the expected return type
            Ok((error_atom, "".to_string(), "".to_string()))
        }
        Err(_) => {
            let error_atom = handle_error("Panic occurred in get_comment")?;
            // Convert to the expected return type
            Ok((error_atom, "".to_string(), "".to_string()))
        }
    }
}

#[rustler::nif]
pub fn update_comment(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
    text: String,
    author: Option<String>,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) {
            // Try to find the comment
            let comments = sheet.get_comments_mut();
            let mut found = false;

            for comment in comments.iter_mut() {
                if comment.get_coordinate().to_string() == cell_address {
                    // Update the comment text
                    comment.set_text_string(text);

                    // Update author if provided
                    if let Some(new_author) = author {
                        comment.set_author(new_author);
                    }

                    found = true;
                    break;
                }
            }

            if found {
                Ok(atoms::ok())
            } else {
                Err(format!("No comment found at cell '{}'", cell_address))
            }
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => handle_error(&msg),
        Err(_) => handle_error("Panic occurred in update_comment"),
    }
}

#[rustler::nif]
pub fn remove_comment(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name_mut(&sheet_name) {
            // Get all comments
            let comments = sheet.get_comments_mut();

            // Find the index of the comment to remove
            let mut index_to_remove = None;
            for (i, comment) in comments.iter().enumerate() {
                if comment.get_coordinate().to_string() == cell_address {
                    index_to_remove = Some(i);
                    break;
                }
            }

            // Remove the comment if found
            if let Some(index) = index_to_remove {
                comments.remove(index);
                Ok(atoms::ok())
            } else {
                Err(format!("No comment found at cell '{}'", cell_address))
            }
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(_)) => Ok(atoms::ok()),
        Ok(Err(msg)) => handle_error(&msg),
        Err(_) => handle_error("Panic occurred in remove_comment"),
    }
}

#[rustler::nif]
pub fn has_comments(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> NifResult<bool> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name(&sheet_name) {
            // Check if the sheet has comments
            Ok(sheet.has_comments())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(result)) => Ok(result),
        Ok(Err(msg)) => {
            handle_error(&msg)?;
            // Return a default value if there's an error
            Ok(false)
        }
        Err(_) => {
            handle_error("Panic occurred in has_comments")?;
            // Return a default value if there's an error
            Ok(false)
        }
    }
}

#[rustler::nif]
pub fn get_comments_count(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> NifResult<usize> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| {
        let spreadsheet = spreadsheet_resource.spreadsheet.lock().unwrap();

        // Get sheet by name
        if let Some(sheet) = spreadsheet.get_sheet_by_name(&sheet_name) {
            // Return the count of comments
            Ok(sheet.get_comments().len())
        } else {
            Err(format!("Sheet '{}' not found", sheet_name))
        }
    }));

    match result {
        Ok(Ok(result)) => Ok(result),
        Ok(Err(msg)) => {
            handle_error(&msg)?;
            // Return a default value if there's an error
            Ok(0)
        }
        Err(_) => {
            handle_error("Panic occurred in get_comments_count")?;
            // Return a default value if there's an error
            Ok(0)
        }
    }
}
