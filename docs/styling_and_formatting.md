# Styling and Formatting

This guide covers the comprehensive styling and formatting capabilities of the UmyaSpreadsheet library for creating professional-looking Excel documents.

## Overview

UmyaSpreadsheet provides a rich set of formatting options:

- Number formats (currency, dates, percentages, etc.)
- Text formatting (fonts, sizes, colors, alignment)
- Cell styling (backgrounds, borders, etc.)
- Row and column formatting
- Alignment and text control

> **Note:** For conditional formatting (data bars, color scales, value-based formatting, etc.), see the dedicated [Conditional Formatting](conditional_formatting.html) guide.

## Cell Formatting

### Number Formats

Number formatting allows you to control how numbers appear in your Excel spreadsheet without changing their underlying values. This is essential for creating professional reports and financial documents.

```elixir
# Set a numeric value
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "1234.5678")

# Apply number formatting
UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "A1", "#,##0.00")

# The value will now appear as "1,234.57" in Excel
```

Common number formats:

| Format Code       | Description                                        | Example    |
| ----------------- | -------------------------------------------------- | ---------- |
| `#,##0.00`        | Number with thousand separators and two decimals   | 1,234.57   |
| `0.00%`           | Percentage with two decimals                       | 42.50%     |
| `$#,##0.00`       | Currency with thousand separators and two decimals | $1,234.57  |
| `[$€-x]#,##0.00`  | Euro currency with thousand separators             | €1,234.57  |
| `yyyy-mm-dd`      | Date in year-month-day format                      | 2023-04-15 |
| `d-mmm-yy`        | Date with abbreviated month                        | 15-Apr-23  |
| `h:mm AM/PM`      | Time with AM/PM                                    | 2:30 PM    |
| `h:mm:ss`         | Time with seconds                                  | 14:30:45   |
| `@`               | Text format (treats numbers as text)               | 1234.5678  |

### Getting Formatted Values

You can retrieve the formatted string representation of a cell as it would appear in Excel. This is useful when you need to display the formatted value in your application:

```elixir
# Set a value with formatting
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "1234.5678")
UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "A1", "#,##0.00")

# Retrieve the formatted value
{:ok, formatted_value} = UmyaSpreadsheet.get_formatted_value(spreadsheet, "Sheet1", "A1")
# => "1,234.57"

# You can still access the raw unformatted value if needed
{:ok, raw_value} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "A1")
# => "1234.5678"
```

### Cell Formatting Inspection

UmyaSpreadsheet provides comprehensive getter functions to inspect all aspects of cell formatting. This is essential for:

- **Style Analysis** - Understanding existing formatting in spreadsheets
- **Conditional Logic** - Making decisions based on current formatting
- **Style Copying** - Programmatically copying formatting between cells
- **Reporting** - Generating reports about document formatting
- **Template Processing** - Analyzing and modifying template documents

#### Font Property Inspection

Retrieve complete font information from any cell:

```elixir
# Get font properties
{:ok, font_name} = UmyaSpreadsheet.get_font_name(spreadsheet, "Sheet1", "A1")
# => "Arial"

{:ok, font_size} = UmyaSpreadsheet.get_font_size(spreadsheet, "Sheet1", "A1")
# => 12.0

{:ok, is_bold} = UmyaSpreadsheet.get_font_bold(spreadsheet, "Sheet1", "A1")
# => true

{:ok, is_italic} = UmyaSpreadsheet.get_font_italic(spreadsheet, "Sheet1", "A1")
# => false

{:ok, underline_style} = UmyaSpreadsheet.get_font_underline(spreadsheet, "Sheet1", "A1")
# => "single"

{:ok, has_strikethrough} = UmyaSpreadsheet.get_font_strikethrough(spreadsheet, "Sheet1", "A1")
# => false

{:ok, font_color} = UmyaSpreadsheet.get_font_color(spreadsheet, "Sheet1", "A1")
# => "FF000000" (black)

{:ok, font_family} = UmyaSpreadsheet.get_font_family(spreadsheet, "Sheet1", "A1")
# => "swiss"

{:ok, font_scheme} = UmyaSpreadsheet.get_font_scheme(spreadsheet, "Sheet1", "A1")
# => "minor"
```

#### Alignment and Text Properties

Inspect text alignment and formatting properties:

```elixir
# Get alignment properties
{:ok, h_align} = UmyaSpreadsheet.get_cell_horizontal_alignment(spreadsheet, "Sheet1", "A1")
# => "center"

{:ok, v_align} = UmyaSpreadsheet.get_cell_vertical_alignment(spreadsheet, "Sheet1", "A1")
# => "middle"

{:ok, wrap_text} = UmyaSpreadsheet.get_cell_wrap_text(spreadsheet, "Sheet1", "A1")
# => true

{:ok, rotation} = UmyaSpreadsheet.get_cell_text_rotation(spreadsheet, "Sheet1", "A1")
# => 45
```

#### Border Inspection

Check border styles and colors for all border positions:

```elixir
# Get border styles
{:ok, top_style} = UmyaSpreadsheet.get_border_style(spreadsheet, "Sheet1", "A1", "top")
# => "thin"

{:ok, bottom_style} = UmyaSpreadsheet.get_border_style(spreadsheet, "Sheet1", "A1", "bottom")
# => "medium"

{:ok, left_style} = UmyaSpreadsheet.get_border_style(spreadsheet, "Sheet1", "A1", "left")
# => "thick"

{:ok, right_style} = UmyaSpreadsheet.get_border_style(spreadsheet, "Sheet1", "A1", "right")
# => "dashed"

# Get border colors
{:ok, top_color} = UmyaSpreadsheet.get_border_color(spreadsheet, "Sheet1", "A1", "top")
# => "FF000000" (black - default when no specific color is set)

{:ok, bottom_color} = UmyaSpreadsheet.get_border_color(spreadsheet, "Sheet1", "A1", "bottom")
# => "FFFF0000" (red)
```

Available border positions: `"top"`, `"bottom"`, `"left"`, `"right"`, `"diagonal"`

#### Fill and Background Inspection

Retrieve cell background and pattern properties:

```elixir
# Get fill properties
{:ok, bg_color} = UmyaSpreadsheet.get_cell_background_color(spreadsheet, "Sheet1", "A1")
# => "FFFFFF00" (yellow)

{:ok, fg_color} = UmyaSpreadsheet.get_cell_foreground_color(spreadsheet, "Sheet1", "A1")
# => "FFFFFFFF" (white)

{:ok, pattern_type} = UmyaSpreadsheet.get_cell_pattern_type(spreadsheet, "Sheet1", "A1")
# => "solid"
```

#### Number Format Inspection

Check number formatting information:

```elixir
# Get number format details
{:ok, format_id} = UmyaSpreadsheet.get_cell_number_format_id(spreadsheet, "Sheet1", "A1")
# => 2

{:ok, format_code} = UmyaSpreadsheet.get_cell_format_code(spreadsheet, "Sheet1", "A1")
# => "#,##0.00"
```

#### Cell Protection Inspection

Check cell protection properties:

```elixir
# Get protection properties
{:ok, is_locked} = UmyaSpreadsheet.get_cell_locked(spreadsheet, "Sheet1", "A1")
# => true

{:ok, is_hidden} = UmyaSpreadsheet.get_cell_hidden(spreadsheet, "Sheet1", "A1")
# => false
```

#### Style Copying Example

Use getter functions to copy formatting between cells:

```elixir
# Copy complete formatting from one cell to another
defmodule StyleCopier do
  def copy_cell_style(spreadsheet, from_sheet, from_cell, to_sheet, to_cell) do
    # Get all font properties
    {:ok, font_name} = UmyaSpreadsheet.get_font_name(spreadsheet, from_sheet, from_cell)
    {:ok, font_size} = UmyaSpreadsheet.get_font_size(spreadsheet, from_sheet, from_cell)
    {:ok, is_bold} = UmyaSpreadsheet.get_font_bold(spreadsheet, from_sheet, from_cell)
    {:ok, is_italic} = UmyaSpreadsheet.get_font_italic(spreadsheet, from_sheet, from_cell)
    {:ok, font_color} = UmyaSpreadsheet.get_font_color(spreadsheet, from_sheet, from_cell)

    # Get alignment properties
    {:ok, h_align} = UmyaSpreadsheet.get_cell_horizontal_alignment(spreadsheet, from_sheet, from_cell)
    {:ok, v_align} = UmyaSpreadsheet.get_cell_vertical_alignment(spreadsheet, from_sheet, from_cell)

    # Get background color
    {:ok, bg_color} = UmyaSpreadsheet.get_cell_background_color(spreadsheet, from_sheet, from_cell)

    # Apply all properties to target cell
    UmyaSpreadsheet.set_font_name(spreadsheet, to_sheet, to_cell, font_name)
    UmyaSpreadsheet.set_font_size(spreadsheet, to_sheet, to_cell, font_size)
    UmyaSpreadsheet.set_font_bold(spreadsheet, to_sheet, to_cell, is_bold)
    UmyaSpreadsheet.set_font_italic(spreadsheet, to_sheet, to_cell, is_italic)
    UmyaSpreadsheet.set_font_color(spreadsheet, to_sheet, to_cell, font_color)
    UmyaSpreadsheet.set_cell_alignment(spreadsheet, to_sheet, to_cell, h_align, v_align)
    UmyaSpreadsheet.set_background_color(spreadsheet, to_sheet, to_cell, bg_color)
  end
end

# Usage
StyleCopier.copy_cell_style(spreadsheet, "Sheet1", "A1", "Sheet1", "B2")
```

> **Error Handling**: All getter functions return `{:ok, value}` tuples on success or `{:error, reason}` tuples when the sheet doesn't exist. Non-existent cells return appropriate default values (e.g., `false` for boolean properties, `"none"` for styles, default colors for color properties).

## Text Formatting

### Font Styling

UmyaSpreadsheet offers comprehensive font styling options to create visually appealing and readable spreadsheets:

```elixir
# Set font bold
UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1", true)

# Set font italic
UmyaSpreadsheet.set_font_italic(spreadsheet, "Sheet1", "A1", true)

# Set font underline with different styles
UmyaSpreadsheet.set_font_underline(spreadsheet, "Sheet1", "A1", "single")  # Single underline
UmyaSpreadsheet.set_font_underline(spreadsheet, "Sheet1", "A2", "double")  # Double underline
UmyaSpreadsheet.set_font_underline(spreadsheet, "Sheet1", "A3", "single_accounting")  # Accounting underline
UmyaSpreadsheet.set_font_underline(spreadsheet, "Sheet1", "A4", "double_accounting")  # Double accounting underline
UmyaSpreadsheet.set_font_underline(spreadsheet, "Sheet1", "A5", "none")  # Remove underline

# Set strikethrough text
UmyaSpreadsheet.set_font_strikethrough(spreadsheet, "Sheet1", "A1", true)

# Set font color
UmyaSpreadsheet.set_font_color(spreadsheet, "Sheet1", "A1", "red")
# Color can be a named color or hex code
UmyaSpreadsheet.set_font_color(spreadsheet, "Sheet1", "A2", "#FF0000")

# Set font size
UmyaSpreadsheet.set_font_size(spreadsheet, "Sheet1", "A1", 14)

# Set font name/family
UmyaSpreadsheet.set_font_name(spreadsheet, "Sheet1", "A1", "Arial")

# Combining multiple font styles
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "Styled Text")
UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A3", true)
UmyaSpreadsheet.set_font_italic(spreadsheet, "Sheet1", "A3", true)
UmyaSpreadsheet.set_font_size(spreadsheet, "Sheet1", "A3", 16)
UmyaSpreadsheet.set_font_color(spreadsheet, "Sheet1", "A3", "blue")
```

### Text Wrapping and Rotation

Control how text is displayed within cells with wrapping and rotation options:

```elixir
# Set multi-line text
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Line 1\nLine 2\nLine 3")

# Enable text wrapping (makes cell display multiple lines)
UmyaSpreadsheet.set_wrap_text(spreadsheet, "Sheet1", "A1", true)

# Set text rotation (value in degrees, 0-180)
UmyaSpreadsheet.set_text_rotation(spreadsheet, "Sheet1", "B1", 45)

# Set vertical text (creates top-to-bottom text)
UmyaSpreadsheet.set_text_rotation(spreadsheet, "Sheet1", "C1", 90)
```

Text wrapping automatically adjusts row height to fit the content, but you may still need to adjust column width manually if text lines are very long.

## Cell and Row Styling

### Background Colors

Apply background colors to make important data stand out or to create visual patterns:

```elixir
# Set cell background color using named color
UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "A1", "yellow")

# Set cell background color using hex code
UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "A2", "#FFFF00")

# Set background color for a range of cells
["A3", "B3", "C3", "D3"]
|> Enum.each(fn cell ->
  UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", cell, "lightblue")
end)
```

### Cell Borders

Add borders to cells to create tables and visual structures:

```elixir
# Add a thin black border around cell A1
UmyaSpreadsheet.set_border(spreadsheet, "Sheet1", "A1", "thin", "black")

# Add a thick red border to the bottom of cell B2
UmyaSpreadsheet.set_border_bottom(spreadsheet, "Sheet1", "B2", "thick", "red")

# Add a medium blue border to the right of cell C3
UmyaSpreadsheet.set_border_right(spreadsheet, "Sheet1", "C3", "medium", "blue")

# Add borders to a range using top, right, bottom, left styles
UmyaSpreadsheet.set_borders(spreadsheet, "Sheet1", "D4:F6", "thin", "thin", "thin", "thin", "black")

# Advanced border styles
UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "A1", "top", "dashed")
UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "A2", "left", "dotted")
UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "A3", "bottom", "double")
UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "A4", "right", "thick")
UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "A5", "all", "medium")
UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "A6", "diagonal", "thin")
```

#### Available Border Styles

| Style | Description |
|-------|-------------|
| `"thin"` | Thin border |
| `"medium"` | Medium border |
| `"thick"` | Thick border |
| `"dashed"` | Dashed border |
| `"dotted"` | Dotted border |
| `"double"` | Double border |
| `"none"` | No border |

### Row Styling

Style an entire row with background and font colors:

```elixir
# Set row 1 with black background and white text
UmyaSpreadsheet.set_row_style(spreadsheet, "Sheet1", 1, "black", "white")

# Set row 3 with light gray background and dark blue text
UmyaSpreadsheet.set_row_style(spreadsheet, "Sheet1", 3, "#EEEEEE", "#000080")
```

### Copying Styles

UmyaSpreadsheet provides efficient ways to copy styles between rows or columns, which can save time when creating consistently formatted spreadsheets:

```elixir
# Copy all styles from row 1 to row 2
UmyaSpreadsheet.copy_row_styling(spreadsheet, "Sheet1", 1, 2)

# Copy styles from row 1 to row 2, but only for columns 3-5
UmyaSpreadsheet.copy_row_styling(spreadsheet, "Sheet1", 1, 2, 3, 5)
```

Similarly, you can copy styles between columns:

```elixir
# Copy all styles from column 1 to column 2
UmyaSpreadsheet.copy_column_styling(spreadsheet, "Sheet1", 1, 2)

# Copy styles from column 1 to column 2, but only for rows 3-5
UmyaSpreadsheet.copy_column_styling(spreadsheet, "Sheet1", 1, 2, 3, 5)
```

This is particularly useful when:

- Creating tables with consistent header and data styling
- Applying a consistent format across multiple sections of a spreadsheet
- Quickly replicating complex styling patterns

## Row and Column Dimensions

Controlling the height of rows and width of columns is essential for creating well-laid-out spreadsheets that are easy to read.

### Row Height

Set the height of rows to accommodate larger text, multi-line content, or to create visual spacing:

```elixir
# Set row 1 height to 30 points (approximately 2x the default height)
UmyaSpreadsheet.set_row_height(spreadsheet, "Sheet1", 1, 30.0)

# Set multiple row heights
[2, 3, 4]
|> Enum.each(fn row ->
  UmyaSpreadsheet.set_row_height(spreadsheet, "Sheet1", row, 20.0)
end)
```

The default row height in Excel is approximately 15 points. Setting a row's height to 0 will hide the row.

### Column Width

Control column width to fit content or create a specific layout:

```elixir
# Set column A width to 15 characters (approximately 2x the default width)
UmyaSpreadsheet.set_column_width(spreadsheet, "Sheet1", "A", 15.0)

# Set column C to be wider for longer text content
UmyaSpreadsheet.set_column_width(spreadsheet, "Sheet1", "C", 25.0)

# Enable auto width for column B (adjusts to fit the content)
UmyaSpreadsheet.set_column_auto_width(spreadsheet, "Sheet1", "B", true)

# Set multiple column widths
["D", "E", "F"]
|> Enum.each(fn col ->
  UmyaSpreadsheet.set_column_width(spreadsheet, "Sheet1", col, 12.0)
end)
```

The default column width in Excel is approximately 8.43 characters. Setting a column's width to 0 will hide the column.

## Cell Alignment

Proper text alignment enhances readability and creates a professional look in your spreadsheets. UmyaSpreadsheet provides complete control over both horizontal and vertical alignment.

### Horizontal and Vertical Alignment

```elixir
# Center text both horizontally and vertically
UmyaSpreadsheet.set_cell_alignment(spreadsheet, "Sheet1", "A1", "center", "center")

# Right-align text at the top of the cell
UmyaSpreadsheet.set_cell_alignment(spreadsheet, "Sheet1", "B2", "right", "top")

# Left-align text at the bottom of the cell
UmyaSpreadsheet.set_cell_alignment(spreadsheet, "Sheet1", "C3", "left", "bottom")

# Justify text with distributed vertical alignment
UmyaSpreadsheet.set_cell_alignment(spreadsheet, "Sheet1", "D4", "justify", "distributed")

# Create a header row with centered, bold text
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Header")
UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1", true)
UmyaSpreadsheet.set_cell_alignment(spreadsheet, "Sheet1", "A1", "center", "center")
```

### Alignment Options

UmyaSpreadsheet provides a comprehensive set of alignment options that match Excel's capabilities:

#### Horizontal Alignment Options

| Option             | Description                                         | Use Case                      |
| ------------------ | --------------------------------------------------- | ----------------------------- |
| `"left"`           | Align text to the left side of the cell             | Default for text              |
| `"center"`         | Center text horizontally in the cell                | Headers, titles               |
| `"right"`          | Align text to the right side of the cell            | Default for numbers           |
| `"justify"`        | Spread text to fill width, aligned at both edges    | Paragraphs, notes             |
| `"distributed"`    | Distribute text evenly across the cell width        | Special formatting            |
| `"fill"`           | Repeat characters to fill the width of the cell     | Pattern creation              |
| `"centerContinuous"` | Center across selection of cells                  | Merged-like appearance        |
| `"general"`        | Use default alignment based on content type         | Mixed content                 |

#### Vertical Alignment Options

| Option           | Description                                         | Use Case                      |
| ---------------- | --------------------------------------------------- | ----------------------------- |
| `"top"`          | Align text to the top of the cell                   | Default for many layouts      |
| `"center"`       | Center text vertically in the cell                  | Centered content, buttons     |
| `"middle"`       | Same as `"center"`                                  | Alternative syntax            |
| `"bottom"`       | Align text to the bottom of the cell                | Default in Excel              |
| `"justify"`      | Distribute text evenly between top and bottom       | Multi-line text               |
| `"distributed"`  | Distribute text with extra padding                  | Special formatting            |

> **Tip:** For the best results with large amounts of text, combine appropriate alignment settings with text wrapping or row height adjustments.

### Text Rotation

UmyaSpreadsheet allows you to rotate text within cells for improved readability in column headers or to fit more information in limited space:

```elixir
# Rotate text 45 degrees clockwise
UmyaSpreadsheet.set_cell_rotation(spreadsheet, "Sheet1", "A1", 45)

# Rotate text 90 degrees for vertical text
UmyaSpreadsheet.set_cell_rotation(spreadsheet, "Sheet1", "B1", 90)

# Rotate text counter-clockwise with negative angles
UmyaSpreadsheet.set_cell_rotation(spreadsheet, "Sheet1", "C1", -45)
```

The rotation angle can be set between -90 and 90 degrees.

### Text Indentation

You can indent text within cells to create hierarchical displays or improve readability:

```elixir
# Basic indentation (1 level)
UmyaSpreadsheet.set_cell_indent(spreadsheet, "Sheet1", "A1", 1)

# Deeper indentation (2 levels)
UmyaSpreadsheet.set_cell_indent(spreadsheet, "Sheet1", "A2", 2)

# Even deeper indentation (5 levels)
UmyaSpreadsheet.set_cell_indent(spreadsheet, "Sheet1", "A3", 5)
```

The indent parameter accepts values from 0 to 255, with each level adding more space before the text.

> **Note:** For more advanced cell formatting options, see the dedicated [Advanced Cell Formatting](advanced_cell_formatting.html) guide.

## Merging Cells

Merging cells is useful for creating headers, titles, or emphasizing specific information across multiple cells:

```elixir
# Merge a range of cells (A1 through C3, creating a 3x3 merged cell)
UmyaSpreadsheet.add_merge_cells(spreadsheet, "Sheet1", "A1:C3")

# Add content to the merged cells (always use the top-left cell reference)
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Merged Cell Content")

# Style the merged cell
UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1", true)
UmyaSpreadsheet.set_font_size(spreadsheet, "Sheet1", "A1", 14)
UmyaSpreadsheet.set_cell_alignment(spreadsheet, "Sheet1", "A1", "center", "center")
```

**Important notes about merged cells:**

- Always use the top-left cell reference (A1 in this example) when setting values or styles
- Any content in other cells within the merged range will be lost
- Merging cells that already contain data will preserve only the top-left cell's content

## Sheet Protection and Visibility

### Protect Sheets

Protect worksheets to prevent accidental or unauthorized changes:

```elixir
# Protect sheet without password
UmyaSpreadsheet.set_sheet_protection(spreadsheet, "Sheet1", nil, true)

# Protect sheet with password
UmyaSpreadsheet.set_sheet_protection(spreadsheet, "Sheet1", "password123", true)
```

### Protect Workbook

Protect the entire workbook structure to prevent users from adding, deleting, or rearranging sheets:

```elixir
UmyaSpreadsheet.set_workbook_protection(spreadsheet, "secure_password")
```

### Sheet Visibility

Control which sheets are visible to users in Excel:

```elixir
# Hide a sheet (can be unhidden by users through Excel UI)
UmyaSpreadsheet.set_sheet_state(spreadsheet, "Sheet2", "hidden")

# Make a sheet very hidden (cannot be unhidden through Excel UI)
UmyaSpreadsheet.set_sheet_state(spreadsheet, "Sheet3", "very_hidden")

# Make a hidden sheet visible again
UmyaSpreadsheet.set_sheet_state(spreadsheet, "Sheet2", "visible")
```

Available sheet states:

- `"visible"` - Sheet is normally visible (default)
- `"hidden"` - Sheet is hidden but can be unhidden by users in Excel
- `"very_hidden"` - Sheet is hidden and cannot be unhidden through the Excel UI

### Hide Gridlines

Control the visibility of gridlines for presentation-quality spreadsheets:

```elixir
# Hide gridlines in Sheet1
UmyaSpreadsheet.set_show_grid_lines(spreadsheet, "Sheet1", false)

# Show gridlines in Sheet1 (default)
UmyaSpreadsheet.set_show_grid_lines(spreadsheet, "Sheet1", true)
```

Hiding gridlines is often useful for:

- Presentation spreadsheets
- Dashboard reports
- Printable forms
- Custom-styled tables where gridlines would be distracting

## Related Documentation

- [Conditional Formatting](conditional_formatting.html) - Add data bars, color scales, and rule-based formatting
- [Data Validation](data_validation.html) - Create dropdown lists and validation rules
- [Sheet Operations](sheet_operations.html) - Comprehensive sheet management functions
- [Image Handling](image_handling.html) - Add and manipulate images in spreadsheets
- [Charts](charts.html) - Create visual data charts

## Best Practices

1. **Consistent Styling** - Define and reuse style patterns for headers, data, and totals
2. **Color Usage** - Use colors sparingly and consistently to highlight important information
3. **Appropriate Formatting** - Choose number formats that match your data type (currency, dates, etc.)
4. **Readability** - Set appropriate column widths and row heights for your content
5. **Protection** - Use sheet protection for production spreadsheets to prevent accidental changes
