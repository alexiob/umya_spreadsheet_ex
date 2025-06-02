# Changelog

## 0.6.16 - 2025-05-31

### Added

- **Conditional Formatting Getter Functions**:
  - **Rule Retrieval Capabilities** - Added comprehensive getter functions for conditional formatting rules
    - New function `UmyaSpreadsheet.ConditionalFormatting.get_conditional_formatting_rules/2-3` retrieves all conditional formatting rules for a sheet or range
    - New function `UmyaSpreadsheet.ConditionalFormatting.get_cell_value_rules/2-3` retrieves cell value conditional formatting rules
    - New function `UmyaSpreadsheet.ConditionalFormatting.get_color_scales/2-3` retrieves color scale conditional formatting rules
    - New function `UmyaSpreadsheet.ConditionalFormatting.get_data_bars/2-3` retrieves data bar conditional formatting rules
    - New function `UmyaSpreadsheet.ConditionalFormatting.get_icon_sets/2-3` retrieves icon set conditional formatting rules
    - New function `UmyaSpreadsheet.ConditionalFormatting.get_top_bottom_rules/2-3` retrieves top/bottom conditional formatting rules
    - New function `UmyaSpreadsheet.ConditionalFormatting.get_above_below_average_rules/2-3` retrieves above/below average conditional formatting rules
    - New function `UmyaSpreadsheet.ConditionalFormatting.get_text_rules/2-3` retrieves text conditional formatting rules

- **VML Drawing Getter Functions**:
  - **Shape Property Retrieval** - Added comprehensive getter functions for VML shapes
    - New function `UmyaSpreadsheet.VmlDrawing.get_shape_style/3` retrieves the CSS style string
    - New function `UmyaSpreadsheet.VmlDrawing.get_shape_type/3` retrieves the shape type (rect, oval, etc.)
    - New function `UmyaSpreadsheet.VmlDrawing.get_shape_filled/3` checks if a shape has fill enabled
    - New function `UmyaSpreadsheet.VmlDrawing.get_shape_fill_color/3` gets the shape's fill color
    - New function `UmyaSpreadsheet.VmlDrawing.get_shape_stroked/3` checks if a shape has stroke enabled
    - New function `UmyaSpreadsheet.VmlDrawing.get_shape_stroke_color/3` gets the shape's stroke color
    - New function `UmyaSpreadsheet.VmlDrawing.get_shape_stroke_weight/3` gets the shape's stroke weight

- **File Format Options Getters**:
  - **Compression Level Information** - Added getter for default compression level
    - New function `UmyaSpreadsheet.FileFormatOptions.get_compression_level/1` returns the default compression level
  - **Encryption Status Check** - Added functions to check encryption settings
    - New function `UmyaSpreadsheet.FileFormatOptions.is_encrypted/1` checks if a spreadsheet has encryption enabled
    - New function `UmyaSpreadsheet.FileFormatOptions.get_encryption_algorithm/1` returns the encryption algorithm used

### Fixed

- **NIF Function Registration**:
  - **Missing NIF Exports** - Added previously missing NIF function exports for file format options
    - Properly registered file format option functions (`write_with_compression`, `write_with_encryption_options`, `to_binary_xlsx`)
    - Fixed issue where these functions were implemented but not properly exposed

### Documentation

- **File Format Options Documentation**:
  - Enhanced documentation for file format options with getter function examples
  - Added information about checking encryption status and retrieving compression settings

### Technical Details

- **Architecture** - Properly aligned Elixir wrapper functions with underlying NIF implementations
- **Consistency** - Ensured all file format options have corresponding getter functions for inspection

## 0.6.15 - 2025-05-31

### Fixed

- **Formula Functions Module Fixes**:
  - **Parameter Type Mismatches** - Fixed functions passing incorrect types to NIFs
    - Modified 10 formula getter functions to properly extract references from Spreadsheet structs
    - Resolved ArgumentError issues in `get_formula_type/3`, `get_shared_index/3`, `get_reference/3`, `get_bx/3`, `get_data_table_2d/3`, `get_data_table_row/3`, `get_input_1deleted/3`, `get_input_2deleted/3`, `get_r1/3`, and `get_r2/3`
  - **Default Value Handling** - Improved null value handling for formula properties
    - Implemented consistent conversion from default values to `nil`
    - Fixed numeric properties to convert `0` to `nil` when appropriate
    - Fixed boolean properties to convert `false` to `nil` when appropriate
    - Fixed string properties to convert empty strings to `nil` when appropriate
  - **Doctest Corrections** - Updated doctests to match actual function behavior
    - Fixed `get_reference/3` doctest expectations to match true function behavior

### Documentation

- **Formula Function Guide Enhancements**:
  - Added new section on working with formula properties
  - Documented new consistent behavior for formula property default values
  - Added examples for formula property getter functions
  - Clarified return value expectations for different property types

### Technical Details

- **Type Handling** - Fixed all remaining functions that were incorrectly passing Spreadsheet structs to NIFs expecting references
- **Parameter Extraction** - Now correctly extracting `reference` field from Spreadsheet structs before passing to NIFs
- **Default Value Standardization** - Consistently treating default values as `nil` for better developer experience and API consistency
- **Test Coverage** - All 600 doctests and 409 regular tests now pass with the new fixes

## 0.6.14 - 2025-05-30

### Documentation

- **Error Handling Documentation Updates** - Fixed all documentation examples to match actual function return types and demonstrate proper error handling patterns
- **Return Type Clarification** - Added explicit notes about which functions return direct values vs. tuples, ensuring documentation accuracy
- **Function Signature Verification** - Updated all guide examples to reflect actual wrapper function behavior and NIF specifications
- **Auto Filter Documentation** - Fixed error handling examples for `has_auto_filter/2` and `get_auto_filter_range/2` functions

### Fixed

- **Documentation Consistency Updates**:
  - **Error Handling Standardization** - Updated all documentation examples to match actual function return types
    - Fixed workbook protection examples to use proper `{:ok, boolean()}` pattern matching instead of direct if statements
    - Updated comment checking examples to handle `boolean() | {:error, atom()}` return types properly
    - Corrected formula function examples to show proper error handling for `get_defined_names/1`
    - Fixed sheet operations examples to show correct tuple destructuring for `get_active_sheet/1`
    - Updated print settings examples to demonstrate proper error handling for getter functions
  - **Return Type Documentation** - Clarified which functions return direct values vs. tuples
    - Added explicit notes about `get_sheet_count/1` returning direct integer values
    - Documented error handling patterns for functions that can fail vs. those that cannot
    - Fixed inconsistent documentation showing wrong assignment patterns
  - **Function Signature Verification** - Ensured all documented examples match actual wrapper function behavior
    - Verified ErrorHandling.standardize_result() usage patterns across different modules
    - Confirmed NIF function specifications align with wrapper implementations
    - Updated examples to reflect actual return type standardization

### Technical Details

- **Documentation Accuracy** - All getter function examples now correctly demonstrate error handling patterns
- **Consistency Improvements** - Standardized error handling examples across all guide documents
- **User Experience** - Developers can now copy-paste documentation examples with confidence they will work correctly

## 0.6.13 - 2025-05-30

### Added

- **Print Settings Getter Functions**:
  - **Page Setup Inspection** - Retrieve current print and page setup settings
    - `UmyaSpreadsheet.get_page_orientation/2` - Get page orientation ("portrait" or "landscape", default: "portrait")
    - `UmyaSpreadsheet.get_paper_size/2` - Get paper size code (default: 1 for Letter)
    - `UmyaSpreadsheet.get_page_scale/2` - Get page scale percentage (default: 100)
    - `UmyaSpreadsheet.get_fit_to_page/2` - Get fit-to-page settings as {width, height} tuple (default: {1, 1})
  - **Margin Inspection** - Retrieve page and header/footer margin settings
    - `UmyaSpreadsheet.get_page_margins/2` - Get page margins as {top, right, bottom, left} tuple in inches (default: {0.75, 0.7, 0.75, 0.7})
    - `UmyaSpreadsheet.get_header_footer_margins/2` - Get header/footer margins as {header, footer} tuple in inches (default: {0.3, 0.3})
  - **Header/Footer Content Inspection** - Retrieve header and footer text
    - `UmyaSpreadsheet.get_header/2` - Get header text with formatting codes (default: empty string)
    - `UmyaSpreadsheet.get_footer/2` - Get footer text with formatting codes (default: empty string)
  - **Print Options Inspection** - Retrieve print centering and area settings
    - `UmyaSpreadsheet.get_print_centered/2` - Get print centering as {horizontal, vertical} boolean tuple (default: {false, false})
    - `UmyaSpreadsheet.get_print_area/2` - Get print area range or nil if not set (default: nil)
    - `UmyaSpreadsheet.get_print_titles/2` - Get print titles as {rows, columns} tuple or nils if not set (default: {nil, nil})
  - **Native Rust Implementation** - All getter functions implemented in native Rust for optimal performance
  - **Error Handling** - Proper error handling for non-existent sheets with appropriate default value returns
  - **Comprehensive Test Coverage** - Full unit test suite with 19 test cases covering all getter/setter functions, default values, and error scenarios
  - **Documentation Updates** - Complete print settings guide with getter function examples and inspection patterns

### Technical Details

- **DefinedName API Integration** - Properly uses `get_address()` method for retrieving print area and print titles from Excel's defined names
- **Type Safety** - All primitive return values properly dereferenced in Rust NIFs using `*` operator for u32, f64, and bool types
- **String Formatting** - Header/footer text properly formatted using Rust's `format!("{:?}")` for consistent string representation
- **Default Value Consistency** - All getter functions return Excel-compatible default values when settings are not explicitly configured

## 0.6.12 - 2025-05-29

### Added

- **Row and Column Property Getter Functions**:
  - **Row Dimension Inspection** - Retrieve row properties and settings
    - `UmyaSpreadsheet.get_row_height/3` - Get row height in points (default: 15.0)
    - `UmyaSpreadsheet.get_row_hidden/3` - Get row hidden state (default: false)
  - **Column Dimension Inspection** - Retrieve column properties and settings
    - `UmyaSpreadsheet.get_column_width/3` - Get column width in characters (default: 8.43)
    - `UmyaSpreadsheet.get_column_auto_width/3` - Get auto-width enabled state (default: false)
    - `UmyaSpreadsheet.get_column_hidden/3` - Get column hidden state (default: false)
  - **Native Rust Implementation** - All getter functions implemented in native Rust for optimal performance
  - **Error Handling** - Proper error handling for non-existent sheets with appropriate default value returns
  - **Auto-width Persistence Fix** - Modified `set_column_auto_width` to set both `auto_width` and `best_fit` fields for Excel compatibility
  - **Comprehensive Test Coverage** - Full unit test suite with 11 test cases covering all getter/setter functions and error scenarios

### Fixed

- **Auto-width Persistence Issue** - Fixed `get_column_auto_width` to return `best_fit` value since that's what actually persists in Excel files
- **Error Handling Standardization** - Unified error handling across all row/column NIFs to return simple error strings instead of nested error tuples
- **Option Handling in Rust NIFs** - Fixed proper handling of cases where row/column dimensions don't exist by returning appropriate default values

## 0.6.11 - 2025-05-29

### Added

- **Row and Column Property Getter Functions**:
  - **Row Dimension Inspection** - Retrieve row properties and settings
    - `UmyaSpreadsheet.get_row_height/3` - Get row height in points (default: 15.0)
    - `UmyaSpreadsheet.get_row_hidden/3` - Get row hidden state (default: false)
  - **Column Dimension Inspection** - Retrieve column properties and settings
    - `UmyaSpreadsheet.get_column_width/3` - Get column width in characters (default: 8.43)
    - `UmyaSpreadsheet.get_column_auto_width/3` - Get auto-width enabled state (default: false)
    - `UmyaSpreadsheet.get_column_hidden/3` - Get column hidden state (default: false)
  - **Native Rust Implementation** - All getter functions implemented in native Rust for optimal performance
  - **Error Handling** - Proper error handling for non-existent sheets with appropriate default value returns
  - **Auto-width Persistence Fix** - Modified `set_column_auto_width` to set both `auto_width` and `best_fit` fields for Excel compatibility
  - **Comprehensive Test Coverage** - Full unit test suite with 11 test cases covering all getter/setter functions and error scenarios

- **Sheet Property Getter Functions**:
  - **Sheet Metadata Inspection** - Comprehensive sheet information retrieval
    - `UmyaSpreadsheet.SheetFunctions.get_sheet_count/1` - Get total number of sheets in spreadsheet
    - `UmyaSpreadsheet.SheetFunctions.get_active_sheet/1` - Get currently active sheet tab index (0-based)
    - `UmyaSpreadsheet.SheetFunctions.get_sheet_state/2` - Get sheet visibility state ("visible", "hidden", "veryhidden")
    - `UmyaSpreadsheet.SheetFunctions.get_sheet_protection/2` - Get detailed sheet protection settings and status
    - `UmyaSpreadsheet.SheetFunctions.get_merge_cells/2` - Get list of merged cell ranges in a sheet
  - **Native Rust Implementation** - All getter functions implemented in native Rust for optimal performance
  - **Error Handling** - Proper error handling for non-existent sheets and invalid operations
  - **Comprehensive Test Coverage** - Full unit test suite with 8 test cases covering success and error scenarios

- **Sheet View Getter Functions**:
  - **Sheet Display Settings** - Retrieve sheet view display settings
    - `UmyaSpreadsheet.get_show_grid_lines/2` - Get whether gridlines are shown (true/false)
    - `UmyaSpreadsheet.get_zoom_scale/2` - Get zoom level percentage (e.g., 100, 150, 75)
    - `UmyaSpreadsheet.get_tab_color/2` - Get tab color in hex format (e.g., "#FF0000")
    - `UmyaSpreadsheet.get_sheet_view/2` - Get view type ("normal", "page_layout", "page_break_preview")
    - `UmyaSpreadsheet.get_selection/2` - Get active cell and selection range as map

- **Workbook View Getter Functions**:
  - **Active Tab Retrieval** - Get the currently active worksheet in a workbook
    - `UmyaSpreadsheet.get_active_tab/1` - Returns the 0-based index of the active tab
  - **Window Position Retrieval** - Get the position and size of the Excel window
    - `UmyaSpreadsheet.get_workbook_window_position/1` - Returns a map with :x, :y, :width, and :height of the workbook window

- **Workbook Protection Getter Functions**:
  - **Protection Status** - Check if a workbook has protection settings enabled
    - `UmyaSpreadsheet.is_workbook_protected/1` - Returns true if the workbook has protection enabled
  - **Protection Details** - Get detailed information about workbook protection settings
    - `UmyaSpreadsheet.get_workbook_protection_details/1` - Returns a map with detailed protection settings

- **Complete Cell Formatting Getter Functions**:
  - **Font Property Inspection API** - Retrieve all font properties from spreadsheet cells
    - `UmyaSpreadsheet.get_font_name/3` - Get font name/typeface (Arial, Times New Roman, etc.)
    - `UmyaSpreadsheet.get_font_size/3` - Get font size in points
    - `UmyaSpreadsheet.get_font_bold/3` - Get bold state (true/false)
    - `UmyaSpreadsheet.get_font_italic/3` - Get italic state (true/false)
    - `UmyaSpreadsheet.get_font_underline/3` - Get underline style (none, single, double, accounting)
    - `UmyaSpreadsheet.get_font_strikethrough/3` - Get strikethrough state (true/false)
    - `UmyaSpreadsheet.get_font_family/3` - Get font family type (roman, swiss, modern, script, decorative)
    - `UmyaSpreadsheet.get_font_scheme/3` - Get font scheme (major, minor, none)
    - `UmyaSpreadsheet.get_font_color/3` - Get font color (hex codes or theme references)
  - **Cell Alignment Inspection** - Retrieve text alignment and formatting properties
    - `UmyaSpreadsheet.get_cell_horizontal_alignment/3` - Get horizontal alignment (left, center, right, justify, etc.)
    - `UmyaSpreadsheet.get_cell_vertical_alignment/3` - Get vertical alignment (top, middle, bottom, etc.)
    - `UmyaSpreadsheet.get_cell_wrap_text/3` - Get text wrap state (true/false)
    - `UmyaSpreadsheet.get_cell_text_rotation/3` - Get text rotation angle (0-359 degrees)
  - **Border Property Inspection** - Retrieve border styles and colors for all positions
    - `UmyaSpreadsheet.get_border_style/4` - Get border style (thin, medium, thick, dashed, dotted, double, none)
    - `UmyaSpreadsheet.get_border_color/4` - Get border color (hex codes with default black FF000000)
    - Support for all border positions: top, bottom, left, right, diagonal
  - **Fill and Background Inspection** - Retrieve cell background and pattern properties
    - `UmyaSpreadsheet.get_cell_background_color/3` - Get cell background color (hex codes)
    - `UmyaSpreadsheet.get_cell_foreground_color/3` - Get pattern foreground color
    - `UmyaSpreadsheet.get_cell_pattern_type/3` - Get fill pattern type (solid, none, etc.)
  - **Number Format Inspection** - Retrieve number formatting information
    - `UmyaSpreadsheet.get_cell_number_format_id/3` - Get number format ID reference
    - `UmyaSpreadsheet.get_cell_format_code/3` - Get format code string (e.g., "#,##0.00")
  - **Cell Protection Inspection** - Retrieve cell protection properties
    - `UmyaSpreadsheet.get_cell_locked/3` - Get cell locked state (true/false)
    - `UmyaSpreadsheet.get_cell_hidden/3` - Get cell hidden state (true/false)
  - **Table Inspection API** - Comprehensive table property getters for Excel table analysis
    - `UmyaSpreadsheet.get_table/4` - Get specific table information by name (structure, range, columns)
    - `UmyaSpreadsheet.get_table_style/4` - Get table style properties (name, column/row highlights, banding)
    - `UmyaSpreadsheet.get_table_columns/4` - Get table column definitions (names, IDs, totals functions)
    - `UmyaSpreadsheet.get_table_totals_row/4` - Get totals row visibility status (true/false)
  - **Background Functions Enhancement** - Added missing background property getters
    - `UmyaSpreadsheet.get_cell_background_color/3` - Get cell background color (hex codes)
    - `UmyaSpreadsheet.get_cell_foreground_color/3` - Get cell foreground/pattern color (hex codes)
    - `UmyaSpreadsheet.get_cell_pattern_type/3` - Get cell fill pattern type (solid, gray125, etc.)
  - **Comprehensive Test Coverage** - 100% test coverage with proper error handling validation
  - **Consistent API Design** - All getter functions follow the same pattern and return appropriate default values

### Fixed

- Fixed border color getter to return default black color (FF000000) instead of empty string when no color is set
- Fixed foreground color getter to correctly return pattern background color, addressing Excel/OpenXML terminology confusion
- Fixed test expectations for border colors - corrected to expect default black color when no explicit colors are set
- Fixed Rust compilation errors in font getter functions caused by improper `Option<&Cell>` handling
- Fixed Color API method calls in `get_font_color` function (corrected `get_argb()` and `get_theme_index()` usage)
- Resolved all cell formatting getter function compilation issues in native Rust code

## 0.6.10 - 2025-05-29

### Added

- **Enhanced Font Management Features**:
  - **Font Family Control** - Set font family types for text styling
    - `UmyaSpreadsheet.set_font_family/4` - Apply font family numbering (roman, swiss, modern, script, decorative)
  - **Font Scheme Support** - Control theme-aware font schemes
    - `UmyaSpreadsheet.set_font_scheme/4` - Apply font schemes (major, minor, none)
  - **Advanced Typography** - Comprehensive font styling capabilities
    - Combine font family, scheme, and name for consistent typography
  - **Complete Font Styling System** - Integrated font management with existing style properties

### Fixed

- Fixed issue where font family and font scheme properties weren't correctly applied to Excel files

## 0.6.9 - 2025-05-28

### Added

- **Complete Document Properties Support**:
  - **Core Document Properties** - Standard metadata fields for Excel documents
    - `UmyaSpreadsheet.DocumentProperties.set_title/2`, `get_title/1` - Document title management
    - `UmyaSpreadsheet.DocumentProperties.set_description/2`, `get_description/1` - Document description
    - `UmyaSpreadsheet.DocumentProperties.set_subject/2`, `get_subject/1` - Document subject
    - `UmyaSpreadsheet.DocumentProperties.set_keywords/2`, `get_keywords/1` - Document keywords
    - `UmyaSpreadsheet.DocumentProperties.set_category/2`, `get_category/1` - Document category
    - `UmyaSpreadsheet.DocumentProperties.set_creator/2`, `get_creator/1` - Document creator
    - `UmyaSpreadsheet.DocumentProperties.set_last_modified_by/2`, `get_last_modified_by/1` - Last editor
    - `UmyaSpreadsheet.DocumentProperties.set_company/2`, `get_company/1` - Company information
    - `UmyaSpreadsheet.DocumentProperties.set_manager/2`, `get_manager/1` - Manager information
    - `UmyaSpreadsheet.DocumentProperties.set_created/2`, `get_created/1` - Creation date
    - `UmyaSpreadsheet.DocumentProperties.set_modified/2`, `get_modified/1` - Last modification date
  - **Custom Document Properties** - User-defined key-value pairs for application-specific metadata
    - `UmyaSpreadsheet.DocumentProperties.set_custom_property_string/3` - Set string properties
    - `UmyaSpreadsheet.DocumentProperties.set_custom_property_number/3` - Set numeric properties
    - `UmyaSpreadsheet.DocumentProperties.set_custom_property_bool/3` - Set boolean properties
    - `UmyaSpreadsheet.DocumentProperties.set_custom_property_date/3` - Set date properties
    - `UmyaSpreadsheet.DocumentProperties.get_custom_property/2` - Retrieve typed properties
    - `UmyaSpreadsheet.DocumentProperties.remove_custom_property/2` - Remove specific properties
    - `UmyaSpreadsheet.DocumentProperties.clear_custom_properties/1` - Remove all custom properties
  - **Property Management Utilities** - Metadata management and retrieval
    - `UmyaSpreadsheet.DocumentProperties.set_properties/2` - Set multiple properties at once
    - `UmyaSpreadsheet.DocumentProperties.get_all_properties/1` - Retrieve all core properties
    - `UmyaSpreadsheet.DocumentProperties.get_custom_property_names/1` - List all custom property names
    - `UmyaSpreadsheet.DocumentProperties.get_custom_properties_count/1` - Count custom properties
    - `UmyaSpreadsheet.DocumentProperties.has_custom_property/2` - Check if property exists
  - **Comprehensive Documentation** - Complete guide with examples and best practices
    - Detailed guide for core and custom document properties
    - Best practices for date handling, Unicode support, and large numbers
    - Error handling patterns and troubleshooting guidance

- **VML Drawing Support**:
  - **VmlDrawing Module** - Full support for Vector Markup Language elements in spreadsheets
    - `UmyaSpreadsheet.VmlDrawing.create_shape/3` - Create VML shapes in worksheets
    - `UmyaSpreadsheet.VmlDrawing.set_shape_type/4` - Set shape type (rect, oval, line, etc.)
    - `UmyaSpreadsheet.VmlDrawing.set_shape_style/4` - Set CSS style for positioning and dimensions
    - `UmyaSpreadsheet.VmlDrawing.set_shape_filled/4` - Control shape fill visibility
    - `UmyaSpreadsheet.VmlDrawing.set_shape_fill_color/4` - Set shape fill color
    - `UmyaSpreadsheet.VmlDrawing.set_shape_stroked/4` - Control shape outline visibility
    - `UmyaSpreadsheet.VmlDrawing.set_shape_stroke_color/4` - Set outline color
    - `UmyaSpreadsheet.VmlDrawing.set_shape_stroke_weight/4` - Set outline thickness
  - **Legacy Drawing Compatibility** - Support for legacy Office drawing objects
  - **Complex Drawing Object Management** - Creation and control of sophisticated drawing elements

### Fixed

- **Document Properties Module Fixes**:
  - Fixed NIF loading error in UmyaNative module for document properties
  - Fixed missing module declaration for VML support in lib.rs
  - Fixed type conversion issues for floating-point and negative numeric properties
  - Improved date handling with timezone-aware comparisons
  - Enhanced error standardization for property not found cases
  - Improved handling of very large numeric values and scientific notation
  - Added comprehensive property value conversion for proper type support

## 0.6.8 - 2025-05-27

### Added

- **Complete Page Breaks Implementation**:
  - **Row Page Break Management** - Full control over manual row page breaks for print layout
    - `UmyaSpreadsheet.PageBreaks.add_row_page_break/4` - Add manual row breaks at specific rows
    - `UmyaSpreadsheet.PageBreaks.remove_row_page_break/3` - Remove specific row breaks
    - `UmyaSpreadsheet.PageBreaks.get_row_page_breaks/2` - Retrieve all row breaks for a worksheet
    - `UmyaSpreadsheet.PageBreaks.clear_row_page_breaks/2` - Remove all row breaks from worksheet
    - `UmyaSpreadsheet.PageBreaks.has_row_page_break/3` - Check if row break exists at specific row
  - **Column Page Break Management** - Full control over manual column page breaks for print layout
    - `UmyaSpreadsheet.PageBreaks.add_column_page_break/4` - Add manual column breaks at specific columns
    - `UmyaSpreadsheet.PageBreaks.remove_column_page_break/3` - Remove specific column breaks
    - `UmyaSpreadsheet.PageBreaks.get_column_page_breaks/2` - Retrieve all column breaks for a worksheet
    - `UmyaSpreadsheet.PageBreaks.clear_column_page_breaks/2` - Remove all column breaks from worksheet
    - `UmyaSpreadsheet.PageBreaks.has_column_page_break/3` - Check if column break exists at specific column
  - **Bulk Operations & Convenience Functions** - Efficient batch operations for page break management
    - `UmyaSpreadsheet.PageBreaks.add_row_page_breaks/3` - Add multiple row breaks in single operation
    - `UmyaSpreadsheet.PageBreaks.add_column_page_breaks/3` - Add multiple column breaks in single operation
    - `UmyaSpreadsheet.PageBreaks.remove_row_page_breaks/3` - Remove multiple row breaks efficiently
    - `UmyaSpreadsheet.PageBreaks.remove_column_page_breaks/3` - Remove multiple column breaks efficiently
    - `UmyaSpreadsheet.PageBreaks.clear_all_page_breaks/2` - Clear both row and column breaks
    - `UmyaSpreadsheet.PageBreaks.get_all_page_breaks/2` - Get both row and column breaks as single result
  - **Comprehensive Page Breaks Documentation**:
    - **Complete Page Breaks Guide** - Comprehensive documentation for print layout control
      - Guide covers manual page break insertion with row and column break management
      - Detailed examples for basic operations, bulk operations, and advanced usage patterns
      - Best practices for print layout design and performance optimization
      - Real-world examples including reports, invoices, and multi-section documents
      - Complete API reference with error handling patterns and troubleshooting guide

- **Complete OLE Objects Documentation**:
  - **Comprehensive OLE Objects Guide** - Complete documentation for working with embedded objects in Excel
    - Guide covers embedding Word documents (.docx), PowerPoint presentations (.pptx), PDF files (.pdf), and text files (.txt)
    - Detailed examples for creating OLE object collections and managing embedded objects
    - Documentation for all supported ProgIDs and file format detection
    - Complete API reference with error handling patterns and best practices
    - Advanced operations including property management and object extraction
  - **Enhanced Documentation Structure** - Added OLE Objects guide to HexDocs navigation
    - Integrated OLE Objects guide into the main documentation index and mix.exs configuration
    - Fixed documentation reference warnings (UmyaNative, missing file links)
    - Consistent formatting and style matching existing documentation guides
    - Cross-references to related guides (limitations, troubleshooting)

- **Complete Rich Text Support**:
  - **Core Rich Text Functionality** - Create and manipulate formatted text within Excel cells
    - `UmyaSpreadsheet.RichText.create/0` - Create new empty rich text objects
    - `UmyaSpreadsheet.RichText.create_from_html/1` - Parse HTML strings into rich text with formatting
    - `UmyaSpreadsheet.RichText.to_html/1` - Generate HTML output from rich text objects
    - `UmyaSpreadsheet.RichText.set_cell_rich_text/4` - Apply rich text formatting to specific cells
    - `UmyaSpreadsheet.RichText.get_cell_rich_text/3` - Retrieve rich text objects from cells
  - **Text Element Management** - Fine-grained control over formatted text segments
    - `UmyaSpreadsheet.RichText.create_text_element/2` - Create individual text elements with custom formatting
    - `UmyaSpreadsheet.RichText.add_text_element/2` - Add text elements to rich text objects
    - `UmyaSpreadsheet.RichText.add_formatted_text/3` - Convenient method to add formatted text directly
    - `UmyaSpreadsheet.RichText.get_elements/1` - Retrieve all text elements from rich text objects
    - `UmyaSpreadsheet.RichText.get_element_text/1` - Extract text content from individual elements
    - `UmyaSpreadsheet.RichText.get_element_font_properties/1` - Get font properties from text elements
  - **Comprehensive Font Formatting** - Support for all Excel text formatting options
    - **Bold, Italic, Underline, Strikethrough** - Standard text styling options
    - **Font Size and Font Name** - Custom typography control
    - **Text Color** - Support for hex colors and named colors
    - **Mixed Formatting** - Multiple formatting styles within a single cell
    - **Idiomatic Elixir API** - Uses atom keys (`:bold`, `:italic`) with automatic conversion to Rust string keys

- **Robust Implementation**:
  - **Resource Management** - Proper Rust resource handling with automatic cleanup
    - `RichTextResource` and `TextElementResource` for memory-safe operations
    - Thread-safe resource management with Mutex protection
    - Automatic resource cleanup when resources go out of scope
  - **Error Handling** - Comprehensive error management throughout the rich text system
    - Graceful handling of invalid HTML input with fallback to plain text
    - Validation of font properties with appropriate error messages
    - Safe resource access with proper error propagation
  - **Performance Optimized** - Efficient implementation for large documents
    - Zero-copy string handling where possible
    - Efficient HTML parsing and generation
    - Optimized font property serialization/deserialization

### Improved

- **Elixir Interface** - Enhanced user experience with idiomatic Elixir patterns
  - **Atom Key Support** - Use idiomatic Elixir atom keys (`:bold`, `:italic`) instead of string keys
  - **Automatic Key Conversion** - Seamless conversion between Elixir atoms and Rust strings
  - **Consistent API** - Rich text functions follow established UmyaSpreadsheet patterns
  - **Type Specifications** - Complete type specs for all rich text functions
- **Documentation** - Comprehensive documentation for rich text functionality
  - **Rich Text Guide** - Created comprehensive 600+ line guide covering all rich text functionality
    - Complete API documentation with function signatures and return values
    - Quick start examples and step-by-step tutorials for common use cases
    - Font properties documentation with practical examples for all formatting options
    - Best practices for performance, color consistency, font selection, and HTML import
    - Error handling patterns and comprehensive troubleshooting guide
    - Four detailed real-world examples: Formatted Headers, Status Indicators, Multi-line Instructions, HTML Import
  - **Rich Text Examples** - Added complete rich text examples to README.md
    - Basic rich text creation and HTML import examples
    - Mixed formatting demonstrations with multiple colors and styles
    - Integration examples showing rich text with other UmyaSpreadsheet features
  - **Function Documentation** - Detailed docstrings with examples for all rich text functions
  - **Type Documentation** - Clear parameter and return value documentation
  - **Updated Documentation Structure** - Enhanced guides.md with rich text guide reference
    - Added rich text to "Styling and Formatting" section in documentation index
    - Complete feature description with capabilities overview

### Fixed

- **Configuration Warning** - Resolved `Application.get_env/3` deprecation warning in module compilation
  - Updated to use compile-time configuration where appropriate
  - Maintained compatibility with standalone scripts and Mix.install scenarios

## 0.6.7 - 2025-05-26

### Added

- **Complete Excel Hyperlinks Functionality**:
  - **Core Hyperlink Management** - Comprehensive hyperlink creation and management system
    - `add_hyperlink/4-6` - Create hyperlinks with web URLs, email addresses, file paths, and internal references
    - `get_hyperlink/3` - Retrieve hyperlink information from specific cells with complete metadata
    - `get_hyperlinks/2` - Get all hyperlinks from worksheets for bulk operations
    - `update_hyperlink/4-6` - Modify existing hyperlinks with new URLs, tooltips, and properties
    - `remove_hyperlink/3` - Delete individual hyperlinks while preserving cell values
    - `remove_all_hyperlinks/2` - Clear all hyperlinks from worksheets efficiently
  - **Advanced Hyperlink Features** - Rich functionality for complex hyperlink scenarios
    - `add_bulk_hyperlinks/3` - Add multiple hyperlinks efficiently with tuple or map formats
    - `has_hyperlink/3` - Check if specific cells contain hyperlinks
    - `has_hyperlinks/2` - Verify if worksheets contain any hyperlinks
    - `count_hyperlinks/2` - Count total hyperlinks in worksheets for management
    - Support for custom tooltips and descriptions for enhanced user experience
    - Integration with cell values for seamless data presentation
  - **Multiple Hyperlink Types** - Complete support for all Excel hyperlink formats
    - **Web URLs** - HTTP/HTTPS links to external websites and web services
    - **Email Addresses** - Mailto links with optional subject and body parameters
    - **File Paths** - Local and network file references (Windows, macOS, Linux)
    - **Internal References** - Links to other worksheets, cell ranges, and named ranges

- **Comprehensive Hyperlinks Documentation**:
  - **Excel Hyperlinks Guide** - Complete 900+ line guide covering all hyperlink functionality
    - Quick start examples and step-by-step tutorials for common use cases
    - Complete API documentation with function signatures and return values
    - Hyperlink types documentation with practical examples for each format
    - Advanced features guide including bulk operations and worksheet-level management
    - Best practices for URL validation, tooltips, organization, and error recovery
    - Error handling patterns with comprehensive recovery strategies
    - Performance considerations for bulk operations and memory management
    - Three detailed real-world examples: Employee Directory, Financial Report Navigation, Project Documentation Hub
  - **Updated Main Documentation** - Enhanced README.md and documentation structure
    - Added hyperlinks to main features list with detailed capabilities
    - Updated documentation index in guides.md with hyperlinks guide reference
    - Enhanced mix.exs documentation structure for proper HexDocs integration
  - **API Documentation** - Complete function documentation with examples and error handling
    - Type specifications for all hyperlink functions with return value patterns
    - Comprehensive parameter documentation with validation requirements
    - Error scenarios documentation with descriptive error messages

### Improved

- **Documentation Structure** - Enhanced organization for better discoverability
  - Added hyperlinks guide to "Advanced Features" section in documentation groups
  - Updated guides index with comprehensive hyperlinks feature description
  - Enhanced cross-references between hyperlinks and related functionality

## 0.6.6 - 2025-05-26

### Added

- **Complete Excel Tables Functionality**:
  - **Core Table Management** - Comprehensive table creation and management system
    - `add_table/8` - Create structured tables with headers, ranges, and totals row options
    - `get_tables/2` - Retrieve all tables from worksheets with complete metadata
    - `remove_table/3` - Delete tables by name with proper cleanup
    - `has_tables?/2` - Check if worksheets contain tables
    - `count_tables/2` - Count tables in worksheets for management operations
  - **Advanced Table Styling** - Professional table appearance with built-in Excel styles
    - `set_table_style/8` - Apply Excel's built-in table styles (Light, Medium, Dark themes)
    - Style customization options: first/last column emphasis, banded rows/columns
    - Support for 60+ built-in table styles (TableStyleLight1-21, TableStyleMedium1-28, TableStyleDark1-11)
    - `remove_table_style/3` - Remove styling while preserving table structure
  - **Dynamic Column Management** - Add and modify table columns programmatically
    - `add_table_column/6` - Add new columns with totals row functions
    - `modify_table_column/7` - Update existing column properties (name, totals function, labels)
    - Support for all Excel totals functions: sum, average, count, countNums, max, min, stdDev, var
  - **Totals Row Control** - Flexible totals row management
    - `set_table_totals_row/4` - Enable/disable totals rows with automatic calculations
    - Automatic function application based on column configurations
    - Dynamic updates when columns are added or modified

- **Comprehensive Documentation Suite for Tables**:
  - **Excel Tables Guide** - Complete 400+ line guide covering all table functionality
    - Step-by-step table creation and management examples
    - Complete styling guide with all available table styles and options
    - Column management patterns and best practices
    - Real-world examples: Employee Management System, Sales Reports with multiple tables
    - Error handling patterns and recovery strategies
    - Performance considerations and optimization tips
  - **Updated Main Documentation** - Enhanced README.md with table examples
    - Added table functionality to feature list with detailed capabilities
    - Complete table workflow example with styling and column management
    - Updated documentation links to include Excel Tables guide
  - **API Documentation** - Comprehensive function documentation with examples
    - Type specifications for all table functions with tuple return patterns
    - Complete parameter documentation with validation rules
    - Error scenarios and return value documentation

- **Robust Error Handling & Validation**:
  - **Consistent Return Format** - All table functions return `{:ok, value}` or `{:error, reason}` tuples
  - **Comprehensive Validation** - Sheet existence, table name uniqueness, range validation
  - **Descriptive Error Messages** - Clear error reasons for debugging and user feedback
  - **Edge Case Handling** - Non-existent sheets, duplicate table names, invalid ranges

- **Complete Test Coverage**:
  - **Unit Tests** - 15 comprehensive tests covering all table functionality
  - **Integration Tests** - Complete workflow tests from creation to deletion
  - **Error Case Testing** - Validation of all error scenarios and edge cases
  - **Tuple Return Validation** - Ensures consistent API patterns across all functions

### Technical Implementation

- **Native Rust Integration** - High-performance table operations via NIF
  - Created comprehensive table.rs module with proper error handling
  - Memory-safe operations with proper Rust borrowing patterns
  - Type-safe conversions between Rust and Elixir data structures
  - Integration with existing umya-spreadsheet table infrastructure
- **Elixir Wrapper Module** - Clean, idiomatic Elixir API
  - Complete UmyaSpreadsheet.Table module with full documentation
  - Proper delegation patterns and error handling
  - Type specifications for all public functions
  - Integration with main UmyaSpreadsheet module
- **Module Registration** - Proper NIF function exports and library integration
  - Updated lib.rs with table module inclusion and function exports
  - Added table function declarations to UmyaNative module
  - Main module integration with aliases and delegations

### Quality Assurance

- **Zero Compilation Warnings** - Clean codebase with no warnings or errors
- **100% Test Pass Rate** - All 15 table tests passing successfully
- **Memory Safety** - Proper resource management and cleanup
- **API Consistency** - Follows established patterns from other modules

## 0.6.5 - 2025-05-25

### Added

- **Comprehensive Documentation Suite**:
  - **Enhanced Main Module Documentation** - Updated `UmyaSpreadsheet` module with comprehensive architecture overview
    - Added ASCII diagram showing core architecture and specialized function modules
    - Documented data flow (NIF Layer → Spreadsheet Reference → Function Modules → Error Handling → Memory Management)
    - Added thread safety and performance characteristics documentation
    - Included compatibility matrix for Excel versions, Elixir/OTP versions, and platforms
    - Added quick start guide with code examples and error handling patterns
  - **Troubleshooting Guide** - Created comprehensive `TROUBLESHOOTING.md` covering:
    - Installation issues (NIF compilation, precompiled binaries, platform-specific problems)
    - Runtime errors (spreadsheet reference errors, sheet not found, invalid cell references)
    - Performance issues (slow file operations, memory usage growth)
    - File format issues (corrupted files, encoding problems)
    - Memory issues (out of memory errors, optimization strategies)
    - Concurrency issues (race conditions, coordination patterns)
    - Platform-specific solutions for macOS, Linux, and Windows
    - Best practices for error handling, resource management, and performance optimization
  - **Limitations and Compatibility Guide** - Created comprehensive `LIMITATIONS.md` covering:
    - File format support matrix (.xlsx ✅, .xlsm ✅, .xls ❌, etc.)
    - Excel feature support levels (basic operations, formatting, advanced features)
    - Platform compatibility matrix (Linux, macOS, Windows with architecture support)
    - Version compatibility (Elixir/OTP, Excel versions, LibreOffice compatibility)
    - Performance limitations with memory usage guidelines and benchmarks
    - Known issues with workarounds and roadmap information

- **Sheet Management Functions**:
  - `rename_sheet` - Renames an existing worksheet with comprehensive validation
    - Validates sheet existence and name conflicts
    - Prevents renaming to empty or whitespace-only names
    - Returns appropriate error messages for invalid operations
    - Full test coverage with 6 test scenarios

- **Advanced Conditional Formatting - Icon Sets**:
  - `add_icon_set` - Adds icon set conditional formatting rules to cell ranges
    - Supports multiple icon styles: "3_traffic_lights", "5_arrows", "4_red_to_black", etc.
    - Flexible threshold configuration (2-5 thresholds per icon set)
    - Multiple threshold types: "min", "max", "number", "percent", "percentile", "formula"
    - Comprehensive validation for threshold count and types
    - Enhanced error handling for invalid inputs and nonexistent sheets

- **Advanced Conditional Formatting - Above/Below Average Rules**:
  - `add_above_below_average_rule` - Adds statistical conditional formatting rules
    - Rule types: "above", "below", "above_equal", "below_equal"
    - Optional standard deviation parameter for advanced statistical rules
    - Customizable format styles for conditional highlighting
    - Support for colors, fonts, and other formatting options

- **Enhanced Test Coverage**:
  - Added 17 comprehensive tests for new conditional formatting features
  - Integration tests demonstrating combined usage of new features
  - Error case testing for invalid inputs and edge scenarios
  - Test file generation for visual verification of formatting rules

### Fixed

- **Function Call Consistency**: Updated test files to use correct API function names
  - Changed `UmyaSpreadsheet.new_file()` to `UmyaSpreadsheet.new()`
  - Changed `UmyaSpreadsheet.write_file()` to `UmyaSpreadsheet.write()`
- **Return Value Handling**: Fixed case clause errors in conditional formatting functions
  - Properly handle `{:ok, :ok}` return patterns from native functions
  - Added comprehensive pattern matching for all possible return values
  - Consistent error handling across all conditional formatting functions

### Improved

- **Native Function Integration**: Enhanced NIF registration and error handling
- **Documentation**: Added comprehensive documentation with examples for all new functions
- **API Consistency**: New functions follow existing patterns and conventions
- **Error Messages**: Improved error reporting with descriptive messages for debugging

## 0.6.4 - 2025-05-24

### Fixed

- **GitHub Actions workflows**: Fixed critical issues in CI/CD pipeline
  - Updated project paths from incorrect "example" to correct "umya_native"
  - Upgraded to modern action versions (actions/checkout@v4, actions/upload-artifact@v4)
  - Fixed cross-compilation setup with proper musl target handling
  - Added RUSTLER_NIF_VERSION environment variable passthrough for consistent builds
- **Cross-compilation configuration**: Created comprehensive Cross.toml setup
  - Added target-specific Docker images for reliable cross-platform builds
  - Configured environment variable passthrough for NIF compilation
  - Added support for all major platforms (Linux, macOS, Windows)
- **Rust build system improvements**:
  - Created build.rs script with platform-specific link arguments
  - Added musl target detection and handling
  - Fixed umya-spreadsheet dependency from incorrect path to proper crates.io version "2.3.0"
  - Enhanced error handling in compilation process
- **Formula functions**: Fixed test failures in defined name functionality
  - Resolved private method access issues with DefinedName.set_name()
  - Updated to use worksheet.add_defined_name() for proper name creation
  - Improved error handling and validation in formula operations
- **CI environment enhancements**:
  - Added UMYA_SPREADSHEET_BUILD=true environment variable
  - Improved Rust toolchain setup and system dependencies
  - Enhanced build reliability across different platforms

### Improved

- Enhanced build reliability and cross-platform compatibility
- Streamlined release process with proper artifact generation
- Better error reporting in CI/CD pipeline
- Improved documentation for build and release processes

## 0.6.3 - 2025-05-24

### Added

- Advanced formula functionality:
  - `set_formula` - Sets a regular formula in a cell
  - `set_array_formula` - Sets an array formula for a range of cells
  - `create_named_range` - Creates a named range reference
  - `create_defined_name` - Creates a defined name with an associated formula
  - `get_defined_names` - Retrieves all defined names in the workbook
- Auto Filter functionality:
  - `set_auto_filter` - Adds filter dropdown buttons to a range of cells
  - `remove_auto_filter` - Removes an auto filter from a worksheet
  - `has_auto_filter` - Checks if a worksheet has an auto filter
  - `get_auto_filter_range` - Gets the range of an existing auto filter
- Enhanced file format options:
  - `write_with_compression` - Controls the compression level (0-9) for XLSX files
  - `write_with_encryption_options` - Enhanced encryption with algorithm selection and security parameters
  - `to_binary_xlsx` - Converts a spreadsheet to binary for web responses or in-memory processing
- Added test coverage for formula, auto filter, and file format functionality:
  - Tests for regular and array formulas
  - Tests for named ranges and defined names
  - Tests for creating, removing, and querying auto filters
  - Tests for auto filter persistence between save and load operations
  - Tests for file compression and binary conversion

## 0.6.2 - 2025-05-23

### Added

- Window settings for workbooks and sheets:
  - Added `set_selection` for setting active cell and selection range in worksheets
  - Added `set_active_tab` to control which tab is active when opening the workbook
  - Added `set_workbook_window_position` to set Excel application window position and size
- Added new test files:
  - Added tests for sheet selection functionality
  - Created comprehensive tests for workbook view settings
  - Updated existing tests to ensure compatibility with new window features
- Added comment management functionality:
  - `add_comment` - Adds a comment to a cell with text and author
  - `get_comment` - Retrieves comment text and author from a cell
  - `update_comment` - Updates an existing comment with new text and optionally a new author
  - `remove_comment` - Removes a comment from a cell
  - `has_comments` - Checks if a worksheet contains any comments
  - `get_comments_count` - Returns the number of comments in a worksheet

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
