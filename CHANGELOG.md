# Changelog

## 0.6.2 - 2023-10-15

### Added

- Window settings for workbooks and sheets:
  - Added `set_selection` for setting active cell and selection range in worksheets
  - Added `set_active_tab` to control which tab is active when opening the workbook
  - Added `set_workbook_window_position` to set Excel application window position and size
- Added new test files:
  - Added tests for sheet selection functionality
  - Created comprehensive tests for workbook view settings
  - Updated existing tests to ensure compatibility with new window features

### Fixed

- Fixed incorrect function name in tests (changed `create_sheet` to `add_sheet`)
- Fixed Rust API usage for accessing workbook view in spreadsheet structure

## 0.6.1 - 2025-05-23

### Added

- Sheet-level features that have been implemented:
  - Tab colors (`set_tab_color`)
  - Sheet views (normal, page layout, page break preview) (`set_sheet_view`)
  - Zoom settings (`set_zoom_scale`, `set_zoom_scale_normal`, `set_zoom_scale_page_layout`, `set_zoom_scale_page_break`)
  - Freeze panes (`freeze_panes`)
  - Split panes (`split_panes`)

### Fixed

- Release builds

## 0.5.6 - 2025-05-21

### Added

- Advanced cell formatting features:
  - Added `set_font_italic` for italic text formatting
  - Added `set_font_underline` with support for various underline styles (single, double, accounting)
  - Added `set_font_strikethrough` for strikethrough text
  - Enhanced borders with `set_border_style` supporting various styles (dotted, dashed, double)
  - Added `set_cell_rotation` for text rotation within cells
  - Added `set_cell_indent` for text indentation
  - Comprehensive documentation and examples for all new formatting features
  - Test suite covering all advanced formatting options

## 0.5.5 - 2025-05-20

### Added

- Concurrent operation capabilities:
  - Added comprehensive test suite for concurrent spreadsheet operations
  - Documented thread safety guidelines and best practices
  - Verified support for multi-threaded access to independent spreadsheets

### Fixed

- Fixed compilation errors in drawing_functions.rs:
  - Consolidated duplicate imports for drawing structs
  - Fixed type specifications in drawing function implementations
  - Improved error handling and return values

### Improved

- Enhanced documentation structure:
  - Created dedicated thread safety guide for concurrent operations
  - Enhanced shapes and drawing documentation with thread safety guidelines
  - Added examples for concurrent operation patterns
  - Added tests for concurrent shape creation and manipulation

## 0.5.3 - 2025-05-20

### Added

- Drawing and shape functionality:
  - `add_shape` - Add various shapes (rectangles, ellipses, triangles, etc.) to worksheets
  - `add_text_box` - Create text boxes with customizable text, background, and border
  - `add_connector` - Draw connector lines between cells
- Comprehensive shape documentation:
  - Added detailed guide for shapes and drawing
  - Example code for diagrams and flowcharts
  - Support for multiple shape types and color specifications

### Improved

- Updated README with drawing and shape examples
- Added shapes and drawing guide to documentation
- Fixed compatibility issues with umya-spreadsheet 2.3.0
- Added test script for drawing functionality

## 0.5.2 - 2025-05-20

### Added

- Pivot table functionality:
  - `add_pivot_table` - Create pivot tables with customizable row, column, and data fields
  - `has_pivot_tables?` - Check if a sheet contains pivot tables
  - `count_pivot_tables` - Count the number of pivot tables on a sheet
  - `refresh_all_pivot_tables` - Update pivot table data
  - `remove_pivot_table` - Remove a specific pivot table
- Comprehensive pivot table documentation:
  - Added detailed guide for pivot table usage
  - Example code for common pivot table operations
  - Full API documentation with typespecs

### Improved

- Updated README with pivot table examples
- Added pivot table guide to documentation navigation
- Implemented tests for all pivot table functionality
- Exposed more functionality from the underlying Rust library

## 0.5.1 - 2025-05-20

### Added

- Data validation functionality:
  - `add_list_validation` - Create dropdown lists in cells
  - `add_number_validation` - Apply number range constraints
  - `add_date_validation` - Apply date constraints
  - `add_text_length_validation` - Restrict text by character count
  - `add_custom_validation` - Apply custom formula-based validation
  - `remove_data_validation` - Remove validation from cells

### Fixed

- Fixed compatibility issues with camelCase and snake_case operator strings in data validation functions
- Updated tests to properly handle the `{:ok, value}` return type from `get_cell_value`
- Fixed deprecated Rustler initialization in lib.rs by removing explicit function listing
- Improved error handling in data validation functions

## 0.5.0 - 2025-05-19

### Added

- Comprehensive documentation for chart types:
  - Added detailed chart types documentation to `add_chart` and `add_chart_with_options` functions
  - Added a "Supported Chart Types" section to `DEVELOPMENT.md` with all 13 chart types
  - Complete reference for all available chart types (`LineChart`, `PieChart`, `BarChart`, etc.)
- Conditional formatting capabilities with several rule types:
  - Cell value rules (greater than, less than, equal to, between, etc.)
  - Top and bottom ranking rules (top N, bottom N%, etc.)
  - Data bars (visual representation of values)
  - Color scales (color gradients based on values)
  - Text-based rules (contains, begins with, ends with, etc.)
- Enhanced documentation structure:
  - Added dedicated guide documents in the `docs` folder
  - Created a guides index for better navigation
  - Integrated guides into ExDocs for improved documentation experience
- Organized guides by topic:
  - Charts guide
  - CSV export and performance guide
  - Image handling guide
  - Sheet operations guide
  - Styling and formatting guide

### Fixed

- Rust code improvements:
  - Fixed deprecated rustler initialization format in `lib.rs`
  - Fixed multiple unused variables in `chart_functions.rs` by prefixing them with underscore
  - Cleaned up Rust warnings for better maintainability

### Improved

- Documentation quality and accessibility:
  - References to guides from `README.md` for better discoverability
  - Updated main module documentation with links to guides
  - Organized documentation groups in `mix.exs` for improved navigation
  - Enhanced in-code documentation for all chart-related functions

## 0.4.0 - 2025-05-18

### Added

- CSV export functionality:
  - `write_csv/3` - Export specific sheets to CSV format
- Light writer functions for better memory usage:
  - `write_light/2` - Write files with less memory overhead
  - `write_with_password_light/3` - Password protection with memory efficiency
- New empty spreadsheet creation:
  - `new_empty/0` - Create spreadsheets without default worksheets
- Image handling:
  - `add_image/4`
  - `download_image/4`
  - `change_image/4`
- Sheet operations:
  - `clone_sheet/3`
  - `remove_sheet/2`
  - `set_sheet_state/3`
  - `set_sheet_protection/4`
- Row and column operations:
  - `insert_new_row/4`
  - `insert_new_column/4`
  - `insert_new_column_by_index/4`
  - `remove_row/4`
  - `remove_column/4`
  - `remove_column_by_index/4`
  - `copy_row_styling/6`
  - `copy_column_styling/6`
- Chart support:
  - `add_chart/9` with multiple chart types
- Cell formatting:
  - `get_formatted_value/3`
  - `set_number_format/4`
  - `set_font_name/4`
  - `set_wrap_text/4`
- Sheet appearance:
  - `set_show_grid_lines/3`
  - `add_merge_cells/3`
- Column dimensions:
  - `set_column_width/4`
  - `set_column_auto_width/4`
- Workbook protection:
  - `set_workbook_protection/2`
- Enhanced documentation:
  - New guide for styling and formatting
  - New guide for sheet operations
  - New guide for chart support
  - New guide for image handling
- Additional test files for new features

### Improved

- Updated README with comprehensive feature list
- Added examples showcasing new functionality
- Enhanced test coverage for all new features

## 0.3.0 - 2025-05-17

### Breaking Changes

- Updated umya-spreadsheet Rust dependency to v2.3.0 from crates.io
- Updated Rustler to v0.36.1
- Changed some API functions to handle new return types (may affect pattern matching)

### Added

- Font styling functions:
  - `set_font_color/4`
  - `set_font_size/4`
  - `set_font_bold/4`
- Password-protected file creation with `write_with_password/3`
- Better examples in documentation
- Comprehensive test suite with proper ExUnit test files
- Added test files for styling and password protection features
- Added integration tests ported from the Rust library's test suite

### Fixed

- Compatibility with Rustler 0.36.1 by handling `{:ok, val}` tuples properly
- Color handling in `set_background_color` and `set_font_color`
- Font size handling with proper float conversion

### Improved

- Converted demo scripts into proper test files
- Added .gitignore entries for test output files
- Enhanced documentation with testing information
- Imported test files from the Rust library for comprehensive testing

## 0.2.0 - Initial Release

- Basic Excel file operations (create, read, write)
- Cell manipulation
- Basic styling with background colors
- Sheet operations
