# UmyaSpreadsheet Documentation

Welcome to the UmyaSpreadsheet documentation. This collection of guides provides comprehensive information on using the UmyaSpreadsheet library to create, read, and manipulate Excel files in Elixir.

## Available Guides

- [**Charts**](charts.html): Learn how to create and customize various chart types

  - Line charts, pie charts, bar charts, scatter plots, and more
  - Customizing chart appearance with styles and options
  - Advanced 3D chart customization

- [**Comments**](comments.html): Working with cell comments

  - Adding comments to cells with author information
  - Retrieving and updating existing comments
  - Removing comments and counting comments in worksheets

- [**Conditional Formatting**](conditional_formatting.html): Apply dynamic formatting based on cell contents

  - Cell value rules (greater than, less than, equal to, etc.)
  - Color scales and data bars for visual data representation
  - Top/bottom ranking and text-based formatting rules

- [**CSV Export & Performance**](csv_export_and_performance.html): Optimize your spreadsheet operations

  - Converting Excel sheets to CSV format
  - Using light writer functions for improved memory usage
  - Performance recommendations for large spreadsheets

- [**Data Validation**](data_validation.html): Control and restrict cell input values

  - Dropdown lists for predefined options
  - Number, date, and text length validation
  - Custom formula-based validation rules
  - Input messages and error alerts

- [**Document Properties**](document_properties.html): Add and manage document metadata

  - Core document properties (title, author, description, etc.)
  - Custom properties for application-specific metadata
  - Reading, writing and removing properties
  - Best practices for date handling and Unicode support

- [**Excel Tables**](excel_tables.html): Create, style, and manage structured Excel tables

  - Creating tables with headers and data ranges
  - Applying built-in table styles and customization options
  - Adding, modifying, and removing table columns dynamically
  - Configuring totals rows with calculation functions
  - Managing table metadata and column properties

- [**Formula Functions**](formula_functions.html): Work with formulas and named references

  - Setting regular cell formulas
  - Creating array formulas that span multiple cells
  - Defining named ranges and defined names
  - Listing and managing defined names

- [**Hyperlinks**](hyperlinks.html): Manage hyperlinks in Excel spreadsheets

  - Adding web URLs, email addresses, file paths, and internal references
  - Getting, updating, and removing hyperlinks from cells
  - Bulk hyperlink operations and worksheet-level management
  - Integration with cell values and advanced error handling

- [**File Format Options**](file_format_options.html): Control compression, encryption, and delivery

  - Adjusting compression levels for Excel files
  - Advanced encryption and security options
  - Generating binary Excel files for web applications
  - Performance optimizations for different scenarios

- [**Auto Filters**](auto_filters.html): Enable Excel's filtering capability

  - Adding filter dropdown buttons to data ranges
  - Checking and removing auto filters
  - Getting filter range information
  - Best practices for using auto filters

- [**Image Handling**](image_handling.html): Working with images in spreadsheets

  - Adding images to cells
  - Downloading images from spreadsheets
  - Changing existing images

- [**OLE Objects**](ole_objects.html): Embed documents and files in Excel spreadsheets

  - Embedding Word documents, PowerPoint presentations, PDF files, and text files
  - Creating and managing OLE object collections
  - Loading objects from files or binary data
  - Extracting embedded objects and managing properties
  - Working with ProgIDs and file format detection

- [**Page Breaks**](page_breaks.html): Control print layout and page breaking

  - Manual page break insertion with row and column controls
  - Managing print layout for consistent document formatting
  - Row breaks and column breaks for precise pagination
  - Bulk operations and advanced page break management

- [**Pivot Tables**](pivot_tables.html): Create and manage data analysis tools

  - Creating pivot tables from data ranges
  - Configuring row, column, and data fields
  - Refreshing pivot table data

- [**Shapes and Drawing**](shapes_and_drawing.html): Add shapes and drawing elements

  - Creating basic shapes (rectangles, ellipses, etc.)
  - Adding text boxes with formatted text
  - Creating connectors between cells
  - Building diagrams and flowcharts

- [**VML Drawings**](vml_drawings.html): Work with Vector Markup Language (VML) shapes

  - Creating VML shapes for legacy compatibility
  - Shape types (rectangles, ovals, lines, rounded rectangles)
  - Style properties for positioning and sizing
  - Fill and stroke properties for appearance customization
  - Best practices for VML shape management

- [**Sheet View Settings**](sheet_view.html): Control individual worksheet appearance and behavior

  - Zoom levels for different Excel view modes (Normal, Page Layout, Page Break)
  - Grid line visibility controls for cleaner presentations
  - Pane freezing and splitting for large dataset navigation
  - Tab colors for visual organization and categorization
  - Cell selection management and user focus control

- [**Window Settings**](window_settings.html): Control Excel viewing experience

  - Setting active sheet selection and focus
  - Configuring which tab is active when opening the workbook
  - Setting the Excel application window position and size

- [**Thread Safety**](thread_safety.html): Using UmyaSpreadsheet in concurrent environments

  - Thread safety guidelines and considerations
  - Best practices for concurrent operations
  - Examples of safe concurrent patterns
  - Avoiding race conditions with shared spreadsheets

- [**Sheet Operations**](sheet_operations.html): Managing worksheets effectively

  - Adding, removing, and cloning sheets
  - Sheet protection and visibility settings
  - Inspecting sheet properties and metadata (count, state, protection details, merged cells)
  - Working with sheet properties

- [**Styling and Formatting**](styling_and_formatting.html): Making your spreadsheets look professional
  - Font styling (size, color, bold)
  - Cell background colors
  - Number formats and cell alignment
  - Borders and other visual elements

- [**Rich Text Formatting**](rich_text.html): Create cells with multiple formatting styles

  - Multiple font styles within a single cell (bold, italic, underline, strikethrough)
  - Custom colors and font sizes for different text segments
  - HTML import and export for complex formatting
  - Comprehensive font property management with idiomatic Elixir API

- [**Advanced Cell Formatting**](advanced_cell_formatting.html): Enhanced visual styling options
  - Rich text styling (italic, strikethrough, multiple underline styles)
  - Advanced border customization (dashed, dotted, double borders)
  - Text rotation and orientation control
  - Cell text indentation for hierarchical displays

- [**Limitations & Compatibility**](limitations.html): Understanding constraints and compatibility
  - File format support matrix and limitations
  - Excel feature support levels and workarounds
  - Platform and version compatibility information
  - Performance limitations and guidelines
  - Known issues and future roadmap

- [**Troubleshooting Guide**](troubleshooting.html): Solutions for common issues and problems
  - Installation and compilation issues
  - Runtime errors and debugging techniques
  - Performance optimization strategies
  - File format and memory management issues
  - Platform-specific solutions and best practices

## Getting Started

For a quick introduction to UmyaSpreadsheet, visit the [README](../readme.html).

For development information, including contribution guidelines and future improvements,
see the [Development Guide](../DEVELOPMENT.html).
