# Sheet Operations

This guide covers comprehensive operations for managing worksheets, rows, columns, and cells in the UmyaSpreadsheet library.

## Overview

UmyaSpreadsheet provides a complete set of functions to:

- Create, copy, and remove worksheets
- Control sheet visibility and protection settings
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

### Set Row Height

Adjust the height of a specific row:

```elixir
# Set row 1 height to 30 points
UmyaSpreadsheet.set_row_height(spreadsheet, "Sheet1", 1, 30.0)
```

Standard Excel row height is 15.0 points.

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
