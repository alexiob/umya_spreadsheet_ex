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

## Inspecting Cell Background Properties

UmyaSpreadsheet provides functions to inspect the background and fill properties of cells:

```elixir
# Get the background color of a cell
{:ok, bg_color} = UmyaSpreadsheet.get_cell_background_color(spreadsheet, "Sheet1", "A1")
# => {:ok, "#FFFFFF"} or {:ok, nil} if no background color is set

# Get the foreground color of a cell (pattern color)
{:ok, fg_color} = UmyaSpreadsheet.get_cell_foreground_color(spreadsheet, "Sheet1", "A1")
# => {:ok, "#000000"} or {:ok, nil} if no foreground color is set

# Get the pattern type of a cell
{:ok, pattern} = UmyaSpreadsheet.get_cell_pattern_type(spreadsheet, "Sheet1", "A1")
# => {:ok, "solid"} or {:ok, nil} if no pattern is set
```

### Available Pattern Types

Common pattern types include:

- `"solid"` - Solid fill
- `"gray125"` - 12.5% gray pattern
- `"gray0625"` - 6.25% gray pattern
- `"horizontal"` - Horizontal lines
- `"vertical"` - Vertical lines
- `"diagonal"` - Diagonal lines
- `"cross"` - Cross pattern
- `"diagonal_cross"` - Diagonal cross pattern

### Practical Background Inspection Example

```elixir
def inspect_cell_background(spreadsheet, sheet_name, cell_address) do
  with {:ok, bg_color} <- UmyaSpreadsheet.get_cell_background_color(spreadsheet, sheet_name, cell_address),
       {:ok, fg_color} <- UmyaSpreadsheet.get_cell_foreground_color(spreadsheet, sheet_name, cell_address),
       {:ok, pattern} <- UmyaSpreadsheet.get_cell_pattern_type(spreadsheet, sheet_name, cell_address) do

    IO.puts("=== Cell Background Details for #{cell_address} ===")
    IO.puts("Background Color: #{bg_color || "None"}")
    IO.puts("Foreground Color: #{fg_color || "None"}")
    IO.puts("Pattern Type: #{pattern || "None"}")

    case {bg_color, fg_color, pattern} do
      {nil, nil, nil} ->
        IO.puts("Cell has no background formatting")

      {bg, nil, "solid"} when not is_nil(bg) ->
        IO.puts("Cell has solid background color: #{bg}")

      {bg, fg, pat} when not is_nil(bg) and not is_nil(fg) ->
        IO.puts("Cell has patterned background - BG: #{bg}, FG: #{fg}, Pattern: #{pat}")

      _ ->
        IO.puts("Cell has partial background formatting")
    end

    {:ok, %{background: bg_color, foreground: fg_color, pattern: pattern}}
  else
    {:error, reason} ->
      IO.puts("Error inspecting cell background: #{reason}")
      {:error, reason}
  end
end

# Usage
inspect_cell_background(spreadsheet, "Sheet1", "A1")
```

### Bulk Background Analysis

```elixir
def analyze_range_backgrounds(spreadsheet, sheet_name, start_cell, end_cell) do
  # This would require implementing a range parser, but shows the concept
  cells = parse_range(start_cell, end_cell)  # Helper function needed

  background_summary =
    Enum.map(cells, fn cell ->
      with {:ok, bg} <- UmyaSpreadsheet.get_cell_background_color(spreadsheet, sheet_name, cell),
           {:ok, fg} <- UmyaSpreadsheet.get_cell_foreground_color(spreadsheet, sheet_name, cell),
           {:ok, pattern} <- UmyaSpreadsheet.get_cell_pattern_type(spreadsheet, sheet_name, cell) do
        %{cell: cell, background: bg, foreground: fg, pattern: pattern}
      else
        _ -> %{cell: cell, background: nil, foreground: nil, pattern: nil}
      end
    end)

  # Group by background properties
  grouped = Enum.group_by(background_summary, fn entry ->
    {entry.background, entry.foreground, entry.pattern}
  end)

  IO.puts("=== Background Analysis for #{start_cell}:#{end_cell} ===")
  Enum.each(grouped, fn {{bg, fg, pattern}, cells} ->
    cell_list = Enum.map(cells, & &1.cell) |> Enum.join(", ")
    IO.puts("#{bg || "No BG"}/#{fg || "No FG"}/#{pattern || "No Pattern"}: #{cell_list}")
  end)

  {:ok, background_summary}
end
```

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
