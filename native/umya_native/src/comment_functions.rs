use crate::atoms;
use crate::UmyaSpreadsheet;
use rustler::{Atom, Error as NifError, NifResult};
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
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<(), String> {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        // Get sheet by name
        let sheet = spreadsheet.get_sheet_by_name_mut(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        // Validate cell address format
        if cell_address.trim().is_empty() {
            return Err("Cell address cannot be empty".to_string());
        }

        // Create a new comment
        let mut comment = Comment::default();

        // Set up the comment with cell address, text and author
        comment.new_comment(cell_address.clone());
        comment.set_text_string(text);
        comment.set_author(author);

        // Add the comment to the worksheet
        sheet.add_comments(comment);

        Ok(())
    }));

    match result {
        Ok(Ok(())) => Ok(atoms::ok()),
        Ok(Err(err_msg)) => Err(NifError::Term(Box::new((atoms::error(), err_msg)))),
        Err(_) => Err(NifError::Term(Box::new((atoms::error(), "Error occurred in add_comment operation".to_string())))),
    }
}

#[rustler::nif]
pub fn get_comment(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<(Atom, String, String)> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<(String, String), String> {
        let spreadsheet = spreadsheet_resource.spreadsheet.lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        // Get sheet by name
        let sheet = spreadsheet.get_sheet_by_name(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        // Validate cell address format
        if cell_address.trim().is_empty() {
            return Err("Cell address cannot be empty".to_string());
        }

        // Get comments as hashmap for easy lookup
        let comments = sheet.get_comments_to_hashmap();

        // Find comment for the specified cell address
        let comment = comments.get(&cell_address)
            .ok_or_else(|| format!("No comment found at cell '{}'", cell_address))?;

        // Return comment text and author (without the nested ok atom)
        Ok((
            comment.get_text().get_text().to_string(),
            comment.get_author().to_string(),
        ))
    }));

    match result {
        Ok(Ok((text, author))) => Ok((atoms::ok(), text, author)),
        Ok(Err(err_msg)) => Err(NifError::Term(Box::new((atoms::error(), err_msg)))),
        Err(_) => Err(NifError::Term(Box::new((atoms::error(), "Error occurred in get_comment operation".to_string())))),
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
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<(), String> {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        // Get sheet by name
        let sheet = spreadsheet.get_sheet_by_name_mut(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        // Validate cell address format
        if cell_address.trim().is_empty() {
            return Err("Cell address cannot be empty".to_string());
        }

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
            Ok(())
        } else {
            Err(format!("No comment found at cell '{}'", cell_address))
        }
    }));

    match result {
        Ok(Ok(())) => Ok(atoms::ok()),
        Ok(Err(err_msg)) => Err(NifError::Term(Box::new((atoms::error(), err_msg)))),
        Err(_) => Err(NifError::Term(Box::new((atoms::error(), "Error occurred in update_comment operation".to_string())))),
    }
}

#[rustler::nif]
pub fn remove_comment(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    cell_address: String,
) -> NifResult<Atom> {
    let result = panic::catch_unwind(AssertUnwindSafe(|| -> Result<(), String> {
        let mut spreadsheet = spreadsheet_resource.spreadsheet.lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        // Get sheet by name
        let sheet = spreadsheet.get_sheet_by_name_mut(&sheet_name)
            .ok_or_else(|| format!("Sheet '{}' not found", sheet_name))?;

        // Validate cell address format
        if cell_address.trim().is_empty() {
            return Err("Cell address cannot be empty".to_string());
        }

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
            Ok(())
        } else {
            Err(format!("No comment found at cell '{}'", cell_address))
        }
    }));

    match result {
        Ok(Ok(())) => Ok(atoms::ok()),
        Ok(Err(err_msg)) => Err(NifError::Term(Box::new((atoms::error(), err_msg)))),
        Err(_) => Err(NifError::Term(Box::new((atoms::error(), "Error occurred in remove_comment operation".to_string())))),
    }
}

#[rustler::nif]
pub fn has_comments(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> bool {
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
        Ok(Ok(result)) => result,
        Ok(Err(_msg)) => {
            // Silent error handling - just return default value without logging
            false
        }
        Err(_) => {
            // Silent error handling - just return default value without logging
            false
        }
    }
}

#[rustler::nif]
pub fn get_comments_count(
    spreadsheet_resource: rustler::ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
) -> usize {
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
        Ok(Ok(result)) => result,
        Ok(Err(_msg)) => {
            // Silent error handling - just return default value without logging
            0
        }
        Err(_) => {
            // Silent error handling - just return default value without logging
            0
        }
    }
}
