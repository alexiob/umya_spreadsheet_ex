use std::path::Path;

/// Checks if a file path exists on the filesystem.
///
/// # Arguments
///
/// * `file_path` - The path to the file (can be relative or absolute)
///
/// # Returns
///
/// An `Option<String>` containing the valid path if it exists, or `None` if it doesn't exist
pub fn find_valid_file_path(file_path: &str) -> Option<String> {
    if Path::new(file_path).exists() {
        Some(file_path.to_string())
    } else {
        None
    }
}
