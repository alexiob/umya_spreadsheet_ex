# Development Notes

## Changes Made

1. Updated Rustler dependency to v0.36.1 in mix.exs and Cargo.toml
2. Updated umya-spreadsheet dependency to v2.3.0 from crates.io (replacing local path)
3. Fixed API differences between local version and crates.io version:
   - Changed `get_sheet_by_name` and `get_sheet_by_name_mut` handling from `Result` to `Option`
   - Updated the Rust NIFs to handle the new return types
4. Added font styling functions to the Elixir API:
   - `set_font_color`
   - `set_font_size`
   - `set_font_bold`
5. Updated the README with better descriptions, feature list, usage examples
6. Added package metadata to mix.exs for better documentation generation
7. Created comprehensive test suite:
   - Converted demo scripts to proper test files
   - Created integration tests based on the Rust library tests
   - Imported test files from Rust library for thorough testing
8. Added CSV export functionality:
   - Implemented `write_csv` function to export sheets as CSV files
   - Added tests for CSV export operations
9. Added light writer functions for better memory usage:
   - `write_light` - Writes spreadsheet using less memory
   - `write_with_password_light` - Writes password-protected spreadsheet using less memory
10. Added empty spreadsheet creation:
    - `new_empty` - Creates a spreadsheet without default worksheets
11. Added cell alignment controls:
    - `set_cell_alignment` - Sets horizontal and vertical alignment for cells
    - Added comprehensive alignment tests
12. Added data validation functionality:
    - `add_list_validation` - Create dropdown lists in cells
    - `add_number_validation` - Apply number range constraints
    - `add_date_validation` - Apply date constraints
    - `add_text_length_validation` - Restrict text by character count
    - `add_custom_validation` - Apply custom formula-based validation
    - `remove_data_validation` - Remove validation from cells
13. Added pivot table functionality:
    - `add_pivot_table` - Create pivot tables from data ranges
    - `has_pivot_tables?` - Check if a sheet contains pivot tables
    - `count_pivot_tables` - Count pivot tables in a sheet
    - `refresh_all_pivot_tables` - Refresh pivot data
    - `remove_pivot_table` - Delete a pivot table
14. Added advanced cell formatting features:
    - `set_font_italic` - Enable/disable italic formatting for text
    - `set_font_underline` - Apply various underline styles (single, double, accounting)
    - `set_font_strikethrough` - Enable/disable strikethrough formatting
    - `set_border_style` - Set border styles for different sides (dotted, dashed, double)
    - `set_cell_rotation` - Set text rotation angle within cells
    - `set_cell_indent` - Control text indentation levels

## Future Improvements

1. Add support for more operations:

   - Rename sheet functionality (rename_sheet)
   - Formula support and calculation
   - Conditional formatting

2. Improve error handling:

   - More specific error types/messages
   - Better validation for inputs
   - Improved handling of corrupted files

3. Add more comprehensive documentation:
   - Add module docs with overview of architecture
   - Document limitations and compatibility issues
   - Create a troubleshooting guide

## Features Available in Rust but Missing in the Elixir Wrapper


1. Advanced Sheet Features

Some sheet-level features that appear missing:

- Tab colors
- Sheet views (normal, page layout, page break preview)
- Zoom settings
- Freeze panes
- Split panes
- Window settings

9. Comment Support

The Rust library has comment functionality (comment.rs, comment.rs)
No corresponding functions in the Elixir wrapper for adding/editing cell comments

10. Advanced Formula Support

While basic cell values are supported, there seems to be no specific handling for:

- Array formulas
- Named ranges
- Defined names

11. Auto Filters

The Rust library has auto filter capability (auto_filter.rs)
The Elixir wrapper doesn't expose this filtering functionality

12. File Format Options

The Rust library likely supports more file format options than what's exposed in the Elixir wrapper
For example, more control over XLSX compression, encryption options, etc.
