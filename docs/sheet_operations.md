# Sheet Operations

This guide covers comprehensive operations for managing worksheets, rows, columns, and cells in the UmyaSpreadsheet library.

## Overview

UmyaSpreadsheet provides a complete set of functions to:

- Create, copy, and remove worksheets
- Control sheet visibility and protection settings
- Inspect sheet properties and retrieve metadata
- Manage rows and columns (insert, remove, resize)
- Apply styles to rows and columns
- Work with merged cells and grid lines

## Basic Sheet Operations

### Adding a Sheet

Add a new worksheet to a spreadsheet:

```elixir
UmyaSpreadsheet.add_sheet(spreadsheet, "NewSheet")
```

### Getting Sheet Names

Retrieve a list of all sheet names in the spreadsheet:

```elixir
sheet_names = UmyaSpreadsheet.get_sheet_names(spreadsheet)
# => ["Sheet1", "NewSheet"]
```

### Cloning a Sheet

Create a copy of an existing sheet with a new name:

```elixir
UmyaSpreadsheet.clone_sheet(spreadsheet, "Sheet1", "Sheet1 Copy")
```

### Removing a Sheet

Remove a worksheet from the spreadsheet:

```elixir
UmyaSpreadsheet.remove_sheet(spreadsheet, "Sheet1 Copy")
```

## Sheet Visibility

Control which sheets are visible to users in Excel or other applications:

```elixir
# Hide a sheet (can be unhidden by user in Excel)
UmyaSpreadsheet.set_sheet_state(spreadsheet, "Sheet2", "hidden")

# Make a sheet very hidden (can only be unhidden programmatically)
UmyaSpreadsheet.set_sheet_state(spreadsheet, "Sheet3", "very_hidden")

# Make a sheet visible
UmyaSpreadsheet.set_sheet_state(spreadsheet, "Sheet2", "visible")
```

Available sheet states:

- `"visible"` - Normal visibility (default)
- `"hidden"` - Hidden but can be unhidden by users in Excel
- `"very_hidden"` - Hidden and cannot be unhidden through the Excel UI

## Sheet Protection

Protect worksheets from editing with optional password protection:

```elixir
# Protect without password
UmyaSpreadsheet.set_sheet_protection(spreadsheet, "Sheet1", nil, true)

# Protect with password
UmyaSpreadsheet.set_sheet_protection(spreadsheet, "Sheet1", "password123", true)

# Remove protection
UmyaSpreadsheet.set_sheet_protection(spreadsheet, "Sheet1", nil, false)
```

The protection settings determine what users can modify while the sheet is protected.

## Sheet Information and Inspection

Retrieve comprehensive information about sheets and their properties using the new getter functions:

### Getting Sheet Count

Get the total number of sheets in a spreadsheet:

```elixir
count = UmyaSpreadsheet.SheetFunctions.get_sheet_count(spreadsheet)
# => 3 (if there are 3 sheets in the workbook)
```

### Getting Active Sheet

Get the index of the currently active sheet:

```elixir
active_index = UmyaSpreadsheet.SheetFunctions.get_active_sheet(spreadsheet)
# => {:ok, 0} (zero-based index, meaning first sheet is active)
```

### Getting Sheet Visibility State

Check the visibility state of a specific sheet:

```elixir
{:ok, state} = UmyaSpreadsheet.SheetFunctions.get_sheet_state(spreadsheet, "Sheet1")
# => {:ok, "visible"} or {:ok, "hidden"} or {:ok, "veryhidden"}
```

This is useful for checking the current state before modifying it or for auditing sheet visibility in your application.

### Getting Sheet Protection Details

Retrieve detailed protection settings for a sheet:

```elixir
{:ok, protection} = UmyaSpreadsheet.SheetFunctions.get_sheet_protection(spreadsheet, "Sheet1")
# => {:ok, %{
#   "protected" => "true",
#   "objects" => "false",
#   "scenarios" => "false",
#   "format_cells" => "true",
#   "format_columns" => "true",
#   "format_rows" => "true",
#   "insert_columns" => "true",
#   "insert_rows" => "true",
#   "insert_hyperlinks" => "true",
#   "delete_columns" => "true",
#   "delete_rows" => "true",
#   "select_locked_cells" => "true",
#   "select_unlocked_cells" => "true",
#   "sort" => "true",
#   "auto_filter" => "true",
#   "pivot_tables" => "true"
# }}
```

### Getting Merged Cells

Get a list of all merged cell ranges in a sheet:

```elixir
{:ok, merged_cells} = UmyaSpreadsheet.SheetFunctions.get_merge_cells(spreadsheet, "Sheet1")
# => {:ok, ["A1:C3", "E5:F6"]} (list of cell range strings)
```

This is particularly useful for:

- Auditing merged cell usage
- Avoiding conflicts when adding new merged cells
- Understanding the layout structure of existing spreadsheets

### Error Handling for Getter Functions

All getter functions include proper error handling for invalid sheet names:

```elixir
case UmyaSpreadsheet.SheetFunctions.get_sheet_state(spreadsheet, "NonExistentSheet") do
  {:ok, state} -> IO.puts("Sheet state: #{state}")
  {:error, reason} -> IO.puts("Error: #{reason}")
end
```

This ensures your application can gracefully handle cases where sheet names might not exist.

## Row Operations

### Insert Rows

Add new rows to a worksheet at a specific position:

```elixir
# Insert 3 new rows starting at row 2
UmyaSpreadsheet.insert_new_row(spreadsheet, "Sheet1", 2, 3)
```

This shifts existing rows down, with all content and formatting preserved.

### Remove Rows

Delete rows from a worksheet at a specific position:

```elixir
# Remove 2 rows starting at row 5
UmyaSpreadsheet.remove_row(spreadsheet, "Sheet1", 5, 2)
```

This shifts remaining rows up, with all content and formatting preserved.

### Row Height

Adjust the height of a specific row:

```elixir
# Set row 1 height to 30 points
UmyaSpreadsheet.set_row_height(spreadsheet, "Sheet1", 1, 30.0)
```

Standard Excel row height is 15.0 points.

Retrieve row properties and settings:

```elixir
# Get the current height of row 1
{:ok, height} = UmyaSpreadsheet.get_row_height(spreadsheet, "Sheet1", 1)
# => {:ok, 30.0}

# Check if row 2 is hidden
{:ok, hidden} = UmyaSpreadsheet.get_row_hidden(spreadsheet, "Sheet1", 2)
# => {:ok, false}
```

These getter functions return default values when the row hasn't been explicitly configured:

- Default row height: 15.0 points
- Default hidden: false

### Row Styling

Apply styling to an entire row:

```elixir
# Set background and font color for a row
UmyaSpreadsheet.set_row_style(spreadsheet, "Sheet1", 1, "black", "white")

# Copy row styling from one row to another
UmyaSpreadsheet.copy_row_styling(spreadsheet, "Sheet1", 1, 2)

# Copy row styling for specific columns (from row 1 to row 2, columns 1-3 only)
UmyaSpreadsheet.copy_row_styling(spreadsheet, "Sheet1", 1, 2, 1, 3)
```

## Column Operations

### Insert Columns

Add new columns to a worksheet at a specific position:

```elixir
# Insert 2 new columns starting at column B (using letter notation)
UmyaSpreadsheet.insert_new_column(spreadsheet, "Sheet1", "B", 2)

# Insert 2 new columns starting at column index 2 (B) (using numeric index)
UmyaSpreadsheet.insert_new_column_by_index(spreadsheet, "Sheet1", 2, 2)
```

This shifts existing columns to the right, with all content and formatting preserved.

### Remove Columns

Delete columns from a worksheet at a specific position:

```elixir
# Remove 1 column starting at column C (using letter notation)
UmyaSpreadsheet.remove_column(spreadsheet, "Sheet1", "C", 1)

# Remove 1 column starting at column index 3 (C) (using numeric index)
UmyaSpreadsheet.remove_column_by_index(spreadsheet, "Sheet1", 3, 1)
```

This shifts remaining columns to the left, with all content and formatting preserved.

### Column Width

Adjust the width of a specific column:

```elixir
# Set column A width to 15 points
UmyaSpreadsheet.set_column_width(spreadsheet, "Sheet1", "A", 15.0)

# Enable auto width for column B (adjusts to content width)
UmyaSpreadsheet.set_column_auto_width(spreadsheet, "Sheet1", "B", true)
```

Standard Excel column width is 8.43 characters (approximately equivalent to 64 pixels).

### Column Width and Visibility Inspection

Retrieve column properties and settings:

```elixir
# Get the current width of column A
{:ok, width} = UmyaSpreadsheet.get_column_width(spreadsheet, "Sheet1", "A")
# => {:ok, 15.0}

# Check if auto-width is enabled for column B
{:ok, auto_width} = UmyaSpreadsheet.get_column_auto_width(spreadsheet, "Sheet1", "B")
# => {:ok, true}

# Check if column C is hidden
{:ok, hidden} = UmyaSpreadsheet.get_column_hidden(spreadsheet, "Sheet1", "C")
# => {:ok, false}
```

These getter functions return default values when the column hasn't been explicitly configured:

- Default column width: 8.43 characters
- Default auto-width: false
- Default hidden: false

### Column Width and Visibility Inspection

Retrieve column properties and settings:

```elixir
# Get the current width of column A
{:ok, width} = UmyaSpreadsheet.get_column_width(spreadsheet, "Sheet1", "A")
# => {:ok, 15.0}

# Check if auto-width is enabled for column B
{:ok, auto_width} = UmyaSpreadsheet.get_column_auto_width(spreadsheet, "Sheet1", "B")
# => {:ok, true}

# Check if column C is hidden
{:ok, hidden} = UmyaSpreadsheet.get_column_hidden(spreadsheet, "Sheet1", "C")
# => {:ok, false}
```

These getter functions return default values when the column hasn't been explicitly configured:

- Default column width: 8.43 characters
- Default auto-width: false
- Default hidden: false

### Column Styling

Copy styling from one column to another:

```elixir
# Copy all styling from column 1 to column 2
UmyaSpreadsheet.copy_column_styling(spreadsheet, "Sheet1", 1, 2)

# Copy column styling for specific rows (from column 1 to column 2, rows 1-5 only)
UmyaSpreadsheet.copy_column_styling(spreadsheet, "Sheet1", 1, 2, 1, 5)
```

## Cell Merging

Merge multiple cells into a single larger cell:

```elixir
# Merge the cells from A1 to C3 (creating a 3x3 merged cell)
UmyaSpreadsheet.add_merge_cells(spreadsheet, "Sheet1", "A1:C3")
```

When cells are merged:

- Only the content in the top-left cell (A1 in this example) is preserved
- Any content in other cells of the range will be lost
- Styling from the top-left cell is applied to the entire merged range

To add content to a merged cell, always use the address of the top-left cell:

```elixir
# Set content for the merged cell
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Merged Cell Content")
```

## Grid Lines

Control the visibility of grid lines in the worksheet:

```elixir
# Hide grid lines in Sheet1
UmyaSpreadsheet.set_show_grid_lines(spreadsheet, "Sheet1", false)

# Show grid lines in Sheet1 (default)
UmyaSpreadsheet.set_show_grid_lines(spreadsheet, "Sheet1", true)
```

Hiding grid lines is often useful for presentation or print layouts. Note that this only affects the visual display and doesn't impact the actual structure of the sheet.

## Related Documentation

- [Styling and Formatting](styling_and_formatting.html) - For cell and worksheet styling options
- [Data Validation](data_validation.html) - For adding validation rules to cells
- [Image Handling](image_handling.html) - For adding images to worksheets
- [Conditional Formatting](conditional_formatting.html) - For rules-based formatting
