# UmyaSpreadsheet

[![CI](https://github.com/alexiob/umya_spreadsheet_ex/actions/workflows/ci.yml/badge.svg)](https://github.com/alexiob/umya_spreadsheet_ex/actions/workflows/ci.yml)
[![Hex.pm](https://img.shields.io/hexpm/v/umya_spreadsheet_ex.svg)](https://hex.pm/packages/umya_spreadsheet_ex)
[![Hex Docs](https://img.shields.io/badge/hex-docs-blue.svg)](https://hexdocs.pm/umya_spreadsheet_ex/)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

![UmyaSpreadsheetEx Logo](/docs/logo.png)

Everything you ever needed (and probably more) to work with Excel files in Elixir, powered by Rust!

An Elixir NIF wrapper for the amazing [umya-spreadsheet](https://github.com/MathNya/umya-spreadsheet) Rust library, providing comprehensive Excel file (.xlsx, .xlsm) manipulation capabilities with the performance benefits of Rust.

## Features

- **Create & Manipulate Spreadsheets**

  - Create new spreadsheets from scratch
  - Read and write existing Excel files
  - Support for both .xlsx and .xlsm file formats
  - CSV export functionality
  - Lightweight writer options for better memory usage

- **Cell Operations**

  - Get and set cell values
  - Retrieve formatted values (numbers, currencies, dates, etc.)
  - Move ranges of cells
  - Merge cells
  - Remove cells
  - Add and manage comments with author info

- **Styling & Formatting**

  - Apply colors (background, font)
  - Set font properties (size, bold, italic, underline, strikethrough, name)
  - Advanced typography with font family and font scheme support
  - Theme-aware font control for consistent document styling
  - Advanced text formatting (various underline styles, text rotation, indentation)
  - Enhanced border styling (dashed, dotted, double, etc.)
  - Apply number formats (currency, percentage, dates, etc.)
  - Enable text wrapping
  - Define column widths and row heights
  - Rich text support with mixed formatting within cells
  - Create and manipulate formatted text elements
  - Convert rich text to/from HTML
  - Conditional formatting (cell value rules, color scales, data bars, top/bottom rules, text rules)
  - Data validation (dropdown lists, number ranges, date constraints, text length limits, custom formulas)

- **Sheet Management**

  - Add, clone, and remove sheets
  - Hide/show sheets
  - Set grid lines visibility
  - Move and reorder sheets
  - Configure window settings and active tab
  - Set cell selection and worksheet view

- **Security**

  - Password protection for workbooks
  - Worksheet level protection
  - Light writer options for password protection

- **Row & Column Operations**

  - Insert and remove rows and columns
  - Adjust row heights and column widths
  - Apply styling to entire rows/columns

- **Print Settings & Page Setup**

  - Control page orientation (portrait/landscape)
  - Configure paper size and scaling
  - Set page margins and headers/footers
  - Define print areas and titles (repeating rows/columns)
  - Control print centering and fit-to-page options

- **Formula Support**
  - Set regular formulas in individual cells
  - Create array formulas across multiple cells
  - Define named ranges for easier formula references
  - Create defined names for constants and formulas
  - List defined names in workbooks

- **Data Organization & Analysis**
  - Add auto filters to column headers for interactive filtering
  - Enable/disable and manage auto filters in worksheets
  - Query auto filter ranges and states
  - Create worksheets with filtered views for easier data analysis

- **Visual Elements**
  - Add images (PNG, JPEG)
  - Create charts (Line, Bar, Pie, and more)
  - Add shapes (rectangles, circles, arrows, etc.)
  - Create text boxes for annotations
  - Connect cells with connector lines
  - Position visual elements precisely

- **Concurrent Operations**
  - Thread-safe creation of multiple spreadsheets
  - Support for parallel read/write operations
  - Guidelines for concurrent spreadsheet manipulation
  - Examples of thread-safe patterns

- **Hyperlinks**
  - Add web URLs, email addresses, file paths, and internal references
  - Get, update, and remove hyperlinks from cells
  - Bulk hyperlink operations for efficient management
  - Integration with cell values and custom tooltips

- **Data Analysis**
  - Create and manage pivot tables
  - Configure row, column, and data fields
  - Refresh pivot table data
  - Position pivot tables on worksheets

- **Excel Tables**
  - Create structured tables with headers and data ranges
  - Apply built-in table styles with customization options
  - Add, modify, and remove table columns dynamically
  - Configure totals rows with various calculation functions
  - Manage table filtering and sorting capabilities
  - Support for table metadata and column properties

## Version Information

This package is built on:

- **umya-spreadsheet**: v2.3.0 - Robust Rust Excel library
- **Rustler**: v0.36.1 - Seamless Rust/Elixir interop with excellent performance

## Documentation

UmyaSpreadsheet has comprehensive guides for all major features:

### Online Documentation

- [**Guides Index**](https://hexdocs.pm/umya_spreadsheet_ex/guides.html) - Starting point for all documentation
- [**Formula Functions**](https://hexdocs.pm/umya_spreadsheet_ex/formula_functions.html) - Working with formulas and named references
- [**Auto Filters**](https://hexdocs.pm/umya_spreadsheet_ex/auto_filters.html) - Creating and managing Excel filter controls
- [**Window Settings**](https://hexdocs.pm/umya_spreadsheet_ex/window_settings.html) - Control how Excel displays your workbooks
- [**Comments**](https://hexdocs.pm/umya_spreadsheet_ex/comments.html) - Adding and managing cell comments
- [**Charts**](https://hexdocs.pm/umya_spreadsheet_ex/charts.html) - Creating and customizing charts
- [**Data Validation**](https://hexdocs.pm/umya_spreadsheet_ex/data_validation.html) - Setting input rules for cells
- [**File Format Options**](https://hexdocs.pm/umya_spreadsheet_ex/file_format_options.html) - Control compression, encryption, and binary format generation

### Reference Documentation

- [**Troubleshooting Guide**](docs/troubleshooting.md) - Solutions for common issues and problems
- [**Limitations & Compatibility**](docs/limitations.md) - Known limitations, compatibility matrix, and workarounds

## Installation

You can install using [igniter](https://hexdocs.pm/igniter) for the most comfortable experience:

```sh
# install igniter if you haven't already
mix archive.install hex igniter_new
# then install umya_spreadsheet_ex
mix igniter.install umya_spreadsheet_ex
```

Or add `umya_spreadsheet_ex` to your list of dependencies in `mix.exs`:

```elixir
def deps do
  [
    {:umya_spreadsheet_ex, "~> 0.6.17"}
  ]
end
```

The package includes precompiled NIF files for common platforms, but will compile from source if needed.

To force NIF compilation, set the `UMYA_SPREADSHEET_BUILD` environment variable to `true`:

```bash
export UMYA_SPREADSHEET_BUILD=true

mix clean;
mix compile;
```

## Quick Start Guide

```elixir
# Create a new spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Write data to cells
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Hello")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "World")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C1", "42")

# Format your data
:ok = UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "A1", "FF0000") # Red
:ok = UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1", true)
:ok = UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "C1", "#,##0.00")

# Save your spreadsheet
:ok = UmyaSpreadsheet.write(spreadsheet, "hello_world.xlsx")

# Read it back
{:ok, loaded_spreadsheet} = UmyaSpreadsheet.read("hello_world.xlsx")
{:ok, value} = UmyaSpreadsheet.get_cell_value(loaded_spreadsheet, "Sheet1", "A1")
# => {:ok, "Hello"}
```

## Documentation and Guides

For more detailed examples and complete API documentation, visit:
[https://hexdocs.pm/umya_spreadsheet_ex](https://hexdocs.pm/umya_spreadsheet_ex)

We provide detailed guides for specific features:

- [**Charts**](https://hexdocs.pm/umya_spreadsheet_ex/charts.html) - Creating and customizing various chart types
- [**CSV Export & Performance**](https://hexdocs.pm/umya_spreadsheet_ex/csv_export_and_performance.html) - CSV export and optimized writers
- [**Data Validation**](https://hexdocs.pm/umya_spreadsheet_ex/data_validation.html) - Control and validate cell input values
- [**Excel Tables**](https://hexdocs.pm/umya_spreadsheet_ex/excel_tables.html) - Create, style, and manage structured Excel tables
- [**File Format Options**](https://hexdocs.pm/umya_spreadsheet_ex/file_format_options.html) - Control compression, encryption, and binary Excel generation
- [**Image Handling**](https://hexdocs.pm/umya_spreadsheet_ex/image_handling.html) - Working with images in spreadsheets
- [**Pivot Tables**](https://hexdocs.pm/umya_spreadsheet_ex/pivot_tables.html) - Create and manage data analysis pivot tables
- [**Print Settings**](https://hexdocs.pm/umya_spreadsheet_ex/print_settings.html) - Configure page setup and print options
- [**Sheet Operations**](https://hexdocs.pm/umya_spreadsheet_ex/sheet_operations.html) - Managing worksheets effectively
- [**Styling and Formatting**](https://hexdocs.pm/umya_spreadsheet_ex/styling_and_formatting.html) - Making your spreadsheets look professional

You can find all guides in our [Guide Index](https://hexdocs.pm/umya_spreadsheet_ex/guides.html).

## Complete Usage Examples

### Styling and Formatting

```elixir
# Create a styled spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Apply various styling options
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Styled Title")
:ok = UmyaSpreadsheet.set_font_size(spreadsheet, "Sheet1", "A1", 16)
:ok = UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1", true)
:ok = UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "A1", "CCFFCC") # Light green

# Add numbers with specific formats
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B3", "1234.56")
:ok = UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "B3", "#,##0.00")

:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B4", "0.42")
:ok = UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "B4", "0.00%")

# Currency formatting
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B5", "9999.99")
:ok = UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "B5", "$#,##0.00")

# Date formatting
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B6", "44500") # Excel date serial
:ok = UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "B6", "yyyy-mm-dd")

# Save the spreadsheet
:ok = UmyaSpreadsheet.write(spreadsheet, "formatted_spreadsheet.xlsx")
```

### Rich Text Formatting

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Create a rich text object
rich_text = UmyaSpreadsheet.RichText.create()

# Add formatted text with different styles
:ok = UmyaSpreadsheet.RichText.add_formatted_text(rich_text, "Bold text", %{bold: true})
:ok = UmyaSpreadsheet.RichText.add_formatted_text(rich_text, " and ", %{})
:ok = UmyaSpreadsheet.RichText.add_formatted_text(rich_text, "italic text", %{italic: true})
:ok = UmyaSpreadsheet.RichText.add_formatted_text(rich_text, " with ", %{})
:ok = UmyaSpreadsheet.RichText.add_formatted_text(rich_text, "colored text", %{color: "#FF0000"})

# Set the rich text to a cell
:ok = UmyaSpreadsheet.RichText.set_cell_rich_text(spreadsheet, "Sheet1", "A1", rich_text)

# Alternative: Create rich text from HTML
html_rich_text = UmyaSpreadsheet.RichText.create_from_html("<b>Bold</b> and <i>italic</i> text")
:ok = UmyaSpreadsheet.RichText.set_cell_rich_text(spreadsheet, "Sheet1", "A2", html_rich_text)

# Create individual text elements for more control
element1 = UmyaSpreadsheet.RichText.create_text_element("Large text", %{size: 18, bold: true})
element2 = UmyaSpreadsheet.RichText.create_text_element(" and small text", %{size: 10})

# Add elements to rich text
rich_text2 = UmyaSpreadsheet.RichText.create()
:ok = UmyaSpreadsheet.RichText.add_text_element(rich_text2, element1)
:ok = UmyaSpreadsheet.RichText.add_text_element(rich_text2, element2)
:ok = UmyaSpreadsheet.RichText.set_cell_rich_text(spreadsheet, "Sheet1", "A3", rich_text2)

# Generate HTML from rich text
html_output = UmyaSpreadsheet.RichText.to_html(rich_text)
# => "<b>Bold text</b> and <i>italic text</i> with <span style=\"color:#FF0000\">colored text</span>"

# Get font properties from text elements
{:ok, props} = UmyaSpreadsheet.RichText.get_element_font_properties(element1)
# props[:bold] => "true"
# props[:size] => "18"

:ok = UmyaSpreadsheet.write(spreadsheet, "rich_text_example.xlsx")
```

### Sheet Operations

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add a new sheet
:ok = UmyaSpreadsheet.add_sheet(spreadsheet, "Data")

# Clone a sheet
:ok = UmyaSpreadsheet.clone_sheet(spreadsheet, "Sheet1", "Sheet1 Copy")

# Get all sheet names
sheet_names = UmyaSpreadsheet.get_sheet_names(spreadsheet)
# => ["Sheet1", "Data", "Sheet1 Copy"]

# Hide a sheet
:ok = UmyaSpreadsheet.set_sheet_state(spreadsheet, "Data", "hidden")

# Remove a sheet
:ok = UmyaSpreadsheet.remove_sheet(spreadsheet, "Sheet1 Copy")

# Turn off grid lines
:ok = UmyaSpreadsheet.set_show_grid_lines(spreadsheet, "Sheet1", false)
```

### Working with Charts

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Prepare some data for the chart
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Month")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Sales")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "January")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", "1200")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "February")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B3", "1500")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A4", "March")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B4", "2200")

# Define data series for the chart
data_series = ["Sheet1!$B$2:$B$4"]
categories = "Sheet1!$A$2:$A$4"
title = "Quarterly Sales"

# Add a column chart
:ok = UmyaSpreadsheet.add_chart(
  spreadsheet,
  "Sheet1",
  "ColumnChart",  # Other options: LineChart, PieChart, BarChart
  "D2",           # Top-left position of chart
  "J10",          # Bottom-right position of chart
  title,          # Chart title
  data_series,    # Data series
  categories      # Categories (optional)
)
```

### Export to CSV

```elixir
# Read an existing spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.read("sales_report.xlsx")

# Export just one sheet to CSV
:ok = UmyaSpreadsheet.write_csv(spreadsheet, "Sheet1", "sheet1_data.csv")
```

### Data Validation

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add a dropdown list to cells
:ok = UmyaSpreadsheet.add_list_validation(
  spreadsheet,
  "Sheet1",
  "A1:A10",
  ["Option 1", "Option 2", "Option 3"],
  true,  # Allow blank values
  "Invalid Selection",  # Error title
  "Please select from the list",  # Error message
  "Selection Required",  # Prompt title
  "Choose an option from the dropdown"  # Prompt message
)

# Add number validation (between 1 and 100)
:ok = UmyaSpreadsheet.add_number_validation(
  spreadsheet,
  "Sheet1",
  "B1:B10",
  "between",  # Operator: between, greaterThan, lessThan, etc.
  1.0,  # Minimum value
  100.0,  # Maximum value
  true,  # Allow blank values
  "Invalid Number",  # Error title
  "Please enter a number between 1 and 100"  # Error message
)

# Add date validation (future dates only)
# Option 1: Using Date struct directly
:ok = UmyaSpreadsheet.add_date_validation(
  spreadsheet,
  "Sheet1",
  "C1:C10",
  "greaterThan",  # Operator
  Date.utc_today(),  # Compare date as Date struct
  nil,  # Second date (for between operator)
  true,  # Allow blank values
  "Invalid Date",  # Error title
  "Please enter a future date"  # Error message
)

# Option 2: Using ISO string format
today_string = Date.utc_today() |> Date.to_iso8601()
:ok = UmyaSpreadsheet.add_date_validation(
  spreadsheet,
  "Sheet1",
  "D1:D10",
  "greaterThan",
  today_string,  # Compare date as string
  nil,
  true,
  "Invalid Date",
  "Please enter a future date"
)

# Add text length validation (max 10 characters)
:ok = UmyaSpreadsheet.add_text_length_validation(
  spreadsheet,
  "Sheet1",
  "D1:D10",
  "lessThanOrEqual",  # Operator
  10,  # Character limit
  nil,  # Second value (for between operator)
  true  # Allow blank values
)

# Remove validation from a range
:ok = UmyaSpreadsheet.remove_data_validation(
  spreadsheet,
  "Sheet1",
  "A5:A10"
)
```

### Improved Memory Usage with Light Writers

```elixir
# Create a large spreadsheet with many sheets
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# [... add lots of data ...]

# Use the light writer for better memory efficiency
:ok = UmyaSpreadsheet.write_light(spreadsheet, "large_file.xlsx")

# Or with password protection
:ok = UmyaSpreadsheet.write_with_password_light(spreadsheet, "secure_large_file.xlsx", "password123")
```

### Working with Pivot Tables

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Create some sample data for the pivot table
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Region")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Product")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C1", "Sales")

:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "North")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", "Apples")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C2", "10000")

:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "North")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B3", "Oranges")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C3", "8000")

:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A4", "South")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B4", "Apples")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C4", "12000")

# Add a sheet for the pivot table
:ok = UmyaSpreadsheet.add_sheet(spreadsheet, "Pivot")

# Create a pivot table
:ok = UmyaSpreadsheet.add_pivot_table(
  spreadsheet,                      # Spreadsheet object
  "Pivot",                          # Destination sheet
  "Sales Analysis",                 # Pivot table name
  "Sheet1",                         # Source sheet
  "A1:C4",                          # Source data range
  "A3",                             # Pivot table top-left position
  [0],                              # Row fields (Region - column index 0)
  [1],                              # Column fields (Product - column index 1)
  [{2, "sum", "Total Sales"}]       # Data fields (Sum of Sales)
)

# Check if a sheet has pivot tables
has_pivots = UmyaSpreadsheet.has_pivot_tables?(spreadsheet, "Pivot")
# => true

# Count pivot tables on a sheet
count = UmyaSpreadsheet.count_pivot_tables(spreadsheet, "Pivot")
# => 1

# Refresh all pivot tables
:ok = UmyaSpreadsheet.refresh_all_pivot_tables(spreadsheet)

# Remove a pivot table
:ok = UmyaSpreadsheet.remove_pivot_table(spreadsheet, "Pivot", "Sales Analysis")
```

### Working with Excel Tables

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add some sample data for the table
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Product")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Category")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C1", "Price")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D1", "Stock")

UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "Laptop")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", "Electronics")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C2", 999.99)
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D2", 50)

UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "Mouse")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B3", "Electronics")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C3", 29.99)
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D3", 100)

# Create a table with the data
{:ok, :ok} = UmyaSpreadsheet.add_table(
  spreadsheet,
  "Sheet1",
  "ProductTable",
  "Product Inventory",
  "A1",
  "D3",
  ["Product", "Category", "Price", "Stock"],
  true  # Show totals row
)

# Apply a table style
{:ok, :ok} = UmyaSpreadsheet.set_table_style(
  spreadsheet,
  "Sheet1",
  "ProductTable",
  "TableStyleMedium9",
  true,   # Show first column
  false,  # Show last column
  true,   # Show banded rows
  false   # Show banded columns
)

# Add a new column to the table
{:ok, :ok} = UmyaSpreadsheet.add_table_column(
  spreadsheet,
  "Sheet1",
  "ProductTable",
  "Total Value",
  "sum",
  "Grand Total"
)

# Check if sheet has tables
{:ok, true} = UmyaSpreadsheet.has_tables?(spreadsheet, "Sheet1")

# Get all tables from the sheet
{:ok, tables} = UmyaSpreadsheet.get_tables(spreadsheet, "Sheet1")
[table | _] = tables
# table["name"] => "ProductTable"
# table["display_name"] => "Product Inventory"

# Remove the table
{:ok, :ok} = UmyaSpreadsheet.remove_table(spreadsheet, "Sheet1", "ProductTable")
```

## Advanced Features

### Row and Column Operations

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Insert 2 new rows at index 3
:ok = UmyaSpreadsheet.insert_new_row(spreadsheet, "Sheet1", 3, 2)

# Insert 1 new column at column C
:ok = UmyaSpreadsheet.insert_new_column(spreadsheet, "Sheet1", "C", 1)

# Remove row at index 5
:ok = UmyaSpreadsheet.remove_row(spreadsheet, "Sheet1", 5, 1)

# Set column width
:ok = UmyaSpreadsheet.set_column_width(spreadsheet, "Sheet1", "A", 15.5)

# Auto-fit column width based on content
:ok = UmyaSpreadsheet.set_column_auto_width(spreadsheet, "Sheet1", "B", true)

# Set row height
:ok = UmyaSpreadsheet.set_row_height(spreadsheet, "Sheet1", 1, 30.0)

# Apply styling to entire row
:ok = UmyaSpreadsheet.set_row_style(spreadsheet, "Sheet1", 1, "EEEEEE", "000000")
```

### Security and Protection

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add data to the spreadsheet
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Confidential Data")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "Employee Salaries")

# Protect the entire workbook
:ok = UmyaSpreadsheet.set_workbook_protection(spreadsheet, true)

# Protect a specific worksheet
:ok = UmyaSpreadsheet.set_sheet_protection(
  spreadsheet,
  "Sheet1",
  true,
  "This sheet is protected"  # Optional message
)

# Save with password protection
:ok = UmyaSpreadsheet.write_with_password(spreadsheet, "confidential.xlsx", "secret123")

# Apply password to an existing file
:ok = UmyaSpreadsheet.set_password("original.xlsx", "protected_copy.xlsx", "secret123")
```

### Print Settings and Page Setup

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add some content to the spreadsheet
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Report Title")
:ok = UmyaSpreadsheet.set_font_size(spreadsheet, "Sheet1", "A1", 16)
:ok = UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1", true)

# Configure print settings
# Set to landscape orientation
:ok = UmyaSpreadsheet.set_page_orientation(spreadsheet, "Sheet1", "landscape")

# Set to A4 paper size
:ok = UmyaSpreadsheet.set_paper_size(spreadsheet, "Sheet1", 9)

# Set page margins (in inches)
:ok = UmyaSpreadsheet.set_page_margins(spreadsheet, "Sheet1", 1.0, 0.75, 1.0, 0.75)

# Set header and footer margins
:ok = UmyaSpreadsheet.set_header_footer_margins(spreadsheet, "Sheet1", 0.5, 0.5)

# Add a custom header
:ok = UmyaSpreadsheet.set_header(spreadsheet, "Sheet1", "&C&\"Arial,Bold\"Confidential Report")

# Add a footer with page numbers
:ok = UmyaSpreadsheet.set_footer(spreadsheet, "Sheet1", "&RPage &P of &N")

# Define a specific print area
:ok = UmyaSpreadsheet.set_print_area(spreadsheet, "Sheet1", "A1:H20")

# Set rows 1-2 to repeat at the top of each printed page
:ok = UmyaSpreadsheet.set_print_titles(spreadsheet, "Sheet1", "1:2", "")

# Center the printout horizontally on the page
:ok = UmyaSpreadsheet.set_print_centered(spreadsheet, "Sheet1", true, false)

# Save the spreadsheet
:ok = UmyaSpreadsheet.write(spreadsheet, "print_ready_report.xlsx")
```

### Working with Images

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add an image to cell A1
:ok = UmyaSpreadsheet.add_image(spreadsheet, "Sheet1", "A1", "path/to/logo.png")

# When reading a spreadsheet with images, you can download them
{:ok, spreadsheet} = UmyaSpreadsheet.read("report_with_images.xlsx")
:ok = UmyaSpreadsheet.download_image(spreadsheet, "Sheet1", "A1", "downloaded_image.png")

# Replace an existing image
:ok = UmyaSpreadsheet.change_image(spreadsheet, "Sheet1", "A1", "path/to/new_logo.png")
```

### Working with Drawing Objects

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add a shape
:ok = UmyaSpreadsheet.add_shape(spreadsheet, "Sheet1", "D3", "rectangle", 200, 100, "blue", "black", 1.0)

# Add a text box
:ok = UmyaSpreadsheet.add_text_box(spreadsheet, "Sheet1", "D5", "Important Note", 200, 100, "yellow", "black", "gray", 1.0)

# Add a connector between cells
:ok = UmyaSpreadsheet.add_connector(spreadsheet, "Sheet1", "A1", "D3", "green", 1.5)
```

## Testing and Development

The library includes comprehensive test cases. To run them:

```bash
mix test
```

Test files are created in the `test/result_files` directory and are automatically ignored by git.

For more details on development, check out the [DEVELOPMENT.md](DEVELOPMENT.md) file.

## Advanced File Format Options

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()
# Add data to the spreadsheet...

# Control the compression level (0-9)
:ok = UmyaSpreadsheet.write_with_compression(spreadsheet, "optimized.xlsx", 8)

# Enhanced encryption with AES256
:ok = UmyaSpreadsheet.write_with_encryption_options(
  spreadsheet,
  "secure.xlsx",
  "myPassword",
  "AES256",        # Algorithm
  "customSaltValue", # Optional salt value
  100000           # Optional spin count
)

# Generate binary XLSX for web responses
xlsx_binary = UmyaSpreadsheet.to_binary_xlsx(spreadsheet)

# In a Phoenix controller:
conn
|> put_resp_content_type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
|> put_resp_header("content-disposition", ~s[attachment; filename="report.xlsx"])
|> send_resp(200, xlsx_binary)
```

## Performance Considerations

- Use `lazy_read/1` for large spreadsheets to load sheets only when accessed
- Use `write_light/2` and `write_with_password_light/3` for better memory efficiency with large files
- For spreadsheets with many sheets, consider using `new_empty/0` and adding only the sheets you need

## License

UmyaSpreadsheetEx is available under the MIT License. See the LICENSE file for more info.

## Acknowledgements

This library is a wrapper around the excellent [umya-spreadsheet](https://github.com/MathNya/umya-spreadsheet) Rust library.
