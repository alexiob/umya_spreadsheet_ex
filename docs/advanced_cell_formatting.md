# Advanced Cell Formatting

UmyaSpreadsheet provides a rich set of features for advanced cell formatting in your Excel documents, including:

- Italic text formatting
- Underline styles (single, double, accounting)
- Strikethrough text
- Border styles (thin, medium, thick, dashed, dotted, double)
- Cell rotation
- Cell indentation

## Italic Formatting

You can make text in a cell italic:

```elixir
UmyaSpreadsheet.set_font_italic(spreadsheet, "Sheet1", "A1", true)
```

## Underline Styles

UmyaSpreadsheet supports multiple underline styles for cell text:

```elixir
# Single underline
UmyaSpreadsheet.set_font_underline(spreadsheet, "Sheet1", "A1", "single")

# Double underline
UmyaSpreadsheet.set_font_underline(spreadsheet, "Sheet1", "A2", "double")

# Single accounting underline
UmyaSpreadsheet.set_font_underline(spreadsheet, "Sheet1", "A3", "single_accounting")

# Double accounting underline
UmyaSpreadsheet.set_font_underline(spreadsheet, "Sheet1", "A4", "double_accounting")

# Remove underline
UmyaSpreadsheet.set_font_underline(spreadsheet, "Sheet1", "A5", "none")
```

## Strikethrough Formatting

You can apply strikethrough formatting to text:

```elixir
UmyaSpreadsheet.set_font_strikethrough(spreadsheet, "Sheet1", "A1", true)
```

## Border Styles

UmyaSpreadsheet allows you to set different border styles for cells:

```elixir
# Apply thin border to all sides
UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "A1", "all", "thin")

# Apply medium border to the left side
UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "A2", "left", "medium")

# Apply thick border to the top side
UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "A3", "top", "thick")

# Apply dashed border to the right side
UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "A4", "right", "dashed")

# Apply dotted border to the bottom side
UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "A5", "bottom", "dotted")

# Apply double border to the diagonal
UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "A6", "diagonal", "double")

# Remove border
UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "A7", "all", "none")
```

Available border positions:
- `"top"` - Top border
- `"right"` - Right border
- `"bottom"` - Bottom border
- `"left"` - Left border
- `"diagonal"` - Diagonal border
- `"all"` - All borders

Available border styles:
- `"thin"` - Thin border
- `"medium"` - Medium border
- `"thick"` - Thick border
- `"dashed"` - Dashed border
- `"dotted"` - Dotted border
- `"double"` - Double border
- `"none"` - No border

## Cell Rotation

You can rotate text within a cell:

```elixir
# Rotate text 45 degrees
UmyaSpreadsheet.set_cell_rotation(spreadsheet, "Sheet1", "A1", 45)

# Rotate text -45 degrees (counter-clockwise)
UmyaSpreadsheet.set_cell_rotation(spreadsheet, "Sheet1", "A2", -45)

# Vertical text (90 degrees)
UmyaSpreadsheet.set_cell_rotation(spreadsheet, "Sheet1", "A3", 90)
```

The angle parameter accepts values from -90 to 90 degrees.

## Cell Indentation

You can set indentation for text within a cell:

```elixir
# Indent text by 1 level
UmyaSpreadsheet.set_cell_indent(spreadsheet, "Sheet1", "A1", 1)

# Indent text by 5 levels
UmyaSpreadsheet.set_cell_indent(spreadsheet, "Sheet1", "A2", 5)
```

The indent parameter accepts values from 0 to 255.

## Combining Formatting Options

You can combine multiple formatting options to create richly formatted cells:

```elixir
# Create a cell with combined formatting
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Formatted Text")
UmyaSpreadsheet.set_font_italic(spreadsheet, "Sheet1", "A1", true)
UmyaSpreadsheet.set_font_underline(spreadsheet, "Sheet1", "A1", "single")
UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "A1", "all", "double")
UmyaSpreadsheet.set_cell_rotation(spreadsheet, "Sheet1", "A1", 15)
```
