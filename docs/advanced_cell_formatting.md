# Advanced Cell Formatting

UmyaSpreadsheet provides a rich set of features for advanced cell formatting in your Excel documents, including:

- Italic text formatting
- Underline styles (single, double, accounting)
- Strikethrough text
- Border styles (thin, medium, thick, dashed, dotted, double)
- Cell rotation
- Cell indentation
- Font family control (Roman, Swiss, Modern, Script, Decorative)
- Font scheme control (Major, Minor, None)

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

## Font Family

You can set the font family type for cells. Excel uses a numbering system to categorize fonts into different families:

```elixir
# Set Roman/Serif font family (e.g., Times New Roman)
UmyaSpreadsheet.set_font_family(spreadsheet, "Sheet1", "A1", "roman")

# Set Swiss/Sans-Serif font family (e.g., Arial, Helvetica)
UmyaSpreadsheet.set_font_family(spreadsheet, "Sheet1", "A2", "swiss")

# Set Modern/Monospace font family (e.g., Courier)
UmyaSpreadsheet.set_font_family(spreadsheet, "Sheet1", "A3", "modern")

# Set Script font family (e.g., Brush Script)
UmyaSpreadsheet.set_font_family(spreadsheet, "Sheet1", "A4", "script")

# Set Decorative font family (e.g., Old English)
UmyaSpreadsheet.set_font_family(spreadsheet, "Sheet1", "A5", "decorative")
```

Available font family types:

- `"roman"` - Roman (Serif) fonts like Times New Roman
- `"swiss"` - Swiss (Sans-Serif) fonts like Arial or Helvetica
- `"modern"` - Modern (Monospace) fonts like Courier
- `"script"` - Script fonts like Brush Script
- `"decorative"` - Decorative fonts like Old English

## Font Scheme

Font schemes are part of Excel's theme system. They allow Excel to substitute appropriate fonts when a workbook is opened on a system that doesn't have the specified font:

```elixir
# Set Major font scheme (typically used for headings)
UmyaSpreadsheet.set_font_scheme(spreadsheet, "Sheet1", "A1", "major")

# Set Minor font scheme (typically used for body text)
UmyaSpreadsheet.set_font_scheme(spreadsheet, "Sheet1", "A2", "minor")

# Set No font scheme (font won't change with theme changes)
UmyaSpreadsheet.set_font_scheme(spreadsheet, "Sheet1", "A3", "none")
```

Available font scheme values:

- `"major"` - Major scheme (typically for headings)
- `"minor"` - Minor scheme (typically for body text)
- `"none"` - No scheme (font won't change with theme changes)

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

### Advanced Typography Example

Here's an example that combines font family, font scheme, and other typography settings:

```elixir
# Create a heading with Major scheme Roman font
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Heading Text")
UmyaSpreadsheet.set_font_name(spreadsheet, "Sheet1", "A1", "Cambria")
UmyaSpreadsheet.set_font_scheme(spreadsheet, "Sheet1", "A1", "major")
UmyaSpreadsheet.set_font_family(spreadsheet, "Sheet1", "A1", "roman")
UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1", true)
UmyaSpreadsheet.set_font_size(spreadsheet, "Sheet1", "A1", 14)

# Create body text with Minor scheme Swiss font
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "Body Text")
UmyaSpreadsheet.set_font_name(spreadsheet, "Sheet1", "A2", "Calibri")
UmyaSpreadsheet.set_font_scheme(spreadsheet, "Sheet1", "A2", "minor")
UmyaSpreadsheet.set_font_family(spreadsheet, "Sheet1", "A2", "swiss")
UmyaSpreadsheet.set_font_size(spreadsheet, "Sheet1", "A2", 11)
```
