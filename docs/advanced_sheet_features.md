# Advanced Sheet Features

This guide explores the advanced sheet-level features available in UmyaSpreadsheet for customizing how Excel sheets are displayed and interacted with.

## Table of Contents

- [Advanced Sheet Features](#advanced-sheet-features)
  - [Table of Contents](#table-of-contents)
  - [Tab Colors](#tab-colors)
  - [Sheet Views](#sheet-views)
  - [Zoom Settings](#zoom-settings)
  - [Freeze Panes](#freeze-panes)
  - [Split Panes](#split-panes)
  - [Setting the Top-Left Cell](#setting-the-top-left-cell)
  - [Showing/Hiding Gridlines](#showinghiding-gridlines)
  - [Combining Features](#combining-features)
  - [Sheet View Getter Functions Reference](#sheet-view-getter-functions-reference)
    - [Tab Color Inspection](#tab-color-inspection)
    - [Sheet View Type Inspection](#sheet-view-type-inspection)
    - [Zoom Scale Inspection](#zoom-scale-inspection)
    - [Grid Lines Visibility Inspection](#grid-lines-visibility-inspection)
    - [Cell Selection Inspection](#cell-selection-inspection)
    - [Practical Examples](#practical-examples)
      - [Checking Sheet Configuration](#checking-sheet-configuration)
      - [Copying Settings Between Sheets](#copying-settings-between-sheets)
      - [Validating Sheet Settings](#validating-sheet-settings)
    - [Default Values](#default-values)
    - [Performance Notes](#performance-notes)

## Tab Colors

You can set the color of worksheet tabs using the `set_tab_color/3` function:

```elixir
# Set the "Sheet1" tab color to red
UmyaSpreadsheet.set_tab_color(spreadsheet, "Sheet1", "#FF0000")

# Set the "Sheet2" tab color to green
UmyaSpreadsheet.set_tab_color(spreadsheet, "Sheet2", "#00FF00")

# Set the "Sheet3" tab color to blue
UmyaSpreadsheet.set_tab_color(spreadsheet, "Sheet3", "#0000FF")
```

## Sheet Views

Excel sheets can be displayed in different view modes. You can set the view mode using the `set_sheet_view/3` function:

```elixir
# Set to normal view
UmyaSpreadsheet.set_sheet_view(spreadsheet, "Sheet1", "normal")

# Set to page layout view
UmyaSpreadsheet.set_sheet_view(spreadsheet, "Sheet1", "page_layout")

# Set to page break preview
UmyaSpreadsheet.set_sheet_view(spreadsheet, "Sheet1", "page_break_preview")
```

## Zoom Settings

You can control the zoom level for different view modes:

```elixir
# Set zoom scale (applies to current view)
UmyaSpreadsheet.set_zoom_scale(spreadsheet, "Sheet1", 150)

# Set zoom scale for normal view
UmyaSpreadsheet.set_zoom_scale_normal(spreadsheet, "Sheet1", 75)

# Set zoom scale for page layout view
UmyaSpreadsheet.set_zoom_scale_page_layout(spreadsheet, "Sheet1", 120)

# Set zoom scale for page break preview
UmyaSpreadsheet.set_zoom_scale_page_break(spreadsheet, "Sheet1", 80)
```

## Freeze Panes

Freeze panes keep rows and/or columns visible when scrolling through a worksheet:

```elixir
# Freeze the top row
UmyaSpreadsheet.freeze_panes(spreadsheet, "Sheet1", 1, 0)

# Freeze the leftmost column
UmyaSpreadsheet.freeze_panes(spreadsheet, "Sheet1", 0, 1)

# Freeze both top row and leftmost column
UmyaSpreadsheet.freeze_panes(spreadsheet, "Sheet1", 1, 1)

# Freeze the top 3 rows and first 2 columns
UmyaSpreadsheet.freeze_panes(spreadsheet, "Sheet1", 3, 2)
```

## Split Panes

Split panes divide the worksheet into separate, independently scrollable regions:

```elixir
# Split the panes at horizontal and vertical position
UmyaSpreadsheet.split_panes(spreadsheet, "Sheet1", 2000, 2000)

# Split just horizontally
UmyaSpreadsheet.split_panes(spreadsheet, "Sheet1", 2000, 0)

# Split just vertically
UmyaSpreadsheet.split_panes(spreadsheet, "Sheet1", 0, 2000)
```

## Setting the Top-Left Cell

You can set which cell will be the top-left visible cell in the sheet:

```elixir
# Set B5 as the top-left visible cell
UmyaSpreadsheet.set_top_left_cell(spreadsheet, "Sheet1", "B5")
```

## Showing/Hiding Gridlines

You can control whether gridlines are shown in a worksheet:

```elixir
# Hide gridlines
UmyaSpreadsheet.set_show_grid_lines(spreadsheet, "Sheet1", false)

# Show gridlines
UmyaSpreadsheet.set_show_grid_lines(spreadsheet, "Sheet1", true)
```

## Combining Features

These features can be combined to create a highly customized worksheet view:

```elixir
# Create a professional financial report view
UmyaSpreadsheet.set_tab_color(spreadsheet, "Financial Report", "#004080")
UmyaSpreadsheet.freeze_panes(spreadsheet, "Financial Report", 2, 1)
UmyaSpreadsheet.set_sheet_view(spreadsheet, "Financial Report", "page_layout")
UmyaSpreadsheet.set_zoom_scale(spreadsheet, "Financial Report", 100)
```

## Sheet View Getter Functions Reference

UmyaSpreadsheet provides comprehensive getter functions to retrieve current sheet view settings. These functions allow you to inspect and verify the current configuration of worksheets.

### Tab Color Inspection

**Function:** `get_tab_color(spreadsheet, sheet_name)`

```elixir
# Get the current tab color
{:ok, color} = UmyaSpreadsheet.get_tab_color(spreadsheet, "Sheet1")

case color do
  "" -> IO.puts("No tab color set (default)")
  hex_color -> IO.puts("Tab color: #{hex_color}")
end
```

**Returns:**

- `{:ok, color_string}` - The hex color code (e.g., "#FF0000") or empty string if no color is set
- `{:error, reason}` - If the sheet doesn't exist or other error

### Sheet View Type Inspection

**Function:** `get_sheet_view(spreadsheet, sheet_name)`

```elixir
# Get the current view mode
{:ok, view_type} = UmyaSpreadsheet.get_sheet_view(spreadsheet, "Sheet1")
IO.puts("Current view: #{view_type}")  # "normal", "page_layout", or "page_break_preview"
```

**Returns:**

- `{:ok, view_type}` - One of "normal", "page_layout", or "page_break_preview"
- `{:error, reason}` - If the sheet doesn't exist or other error

### Zoom Scale Inspection

**Function:** `get_zoom_scale(spreadsheet, sheet_name)`

```elixir
# Get the current zoom level
{:ok, zoom} = UmyaSpreadsheet.get_zoom_scale(spreadsheet, "Sheet1")
IO.puts("Current zoom: #{zoom}%")  # Default: 100
```

**Returns:**

- `{:ok, zoom_percentage}` - Integer zoom percentage (10-400, default: 100)
- `{:error, reason}` - If the sheet doesn't exist or other error

### Grid Lines Visibility Inspection

**Function:** `get_show_grid_lines(spreadsheet, sheet_name)`

```elixir
# Check if grid lines are visible
{:ok, grid_lines_visible} = UmyaSpreadsheet.get_show_grid_lines(spreadsheet, "Sheet1")
IO.puts("Grid lines visible: #{grid_lines_visible}")  # Default: true
```

**Returns:**

- `{:ok, boolean}` - `true` if grid lines are visible, `false` if hidden
- `{:error, reason}` - If the sheet doesn't exist or other error

### Cell Selection Inspection

**Function:** `get_selection(spreadsheet, sheet_name)`

```elixir
# Get the current cell selection
{:ok, selection} = UmyaSpreadsheet.get_selection(spreadsheet, "Sheet1")
IO.inspect(selection)
# %{"active_cell" => "A1", "sqref" => "A1:B5"}
```

**Returns:**

- `{:ok, selection_map}` - Map with "active_cell" and "sqref" keys
  - `"active_cell"` - The currently active cell (e.g., "A1")
  - `"sqref"` - The selected range (e.g., "A1:B5" for ranges, "A1" for single cell)
- `{:error, reason}` - If the sheet doesn't exist or other error

### Practical Examples

#### Checking Sheet Configuration

```elixir
defmodule SheetInspector do
  def inspect_sheet_settings(spreadsheet, sheet_name) do
    IO.puts("=== Sheet Settings for #{sheet_name} ===")

    # Check tab color
    case UmyaSpreadsheet.get_tab_color(spreadsheet, sheet_name) do
      {:ok, ""} -> IO.puts("Tab Color: Default (no color)")
      {:ok, color} -> IO.puts("Tab Color: #{color}")
      {:error, reason} -> IO.puts("Error getting tab color: #{reason}")
    end

    # Check view type
    case UmyaSpreadsheet.get_sheet_view(spreadsheet, sheet_name) do
      {:ok, view} -> IO.puts("View Type: #{view}")
      {:error, reason} -> IO.puts("Error getting view: #{reason}")
    end

    # Check zoom level
    case UmyaSpreadsheet.get_zoom_scale(spreadsheet, sheet_name) do
      {:ok, zoom} -> IO.puts("Zoom Scale: #{zoom}%")
      {:error, reason} -> IO.puts("Error getting zoom: #{reason}")
    end

    # Check grid lines
    case UmyaSpreadsheet.get_show_grid_lines(spreadsheet, sheet_name) do
      {:ok, visible} -> IO.puts("Grid Lines: #{if visible, do: "Visible", else: "Hidden"}")
      {:error, reason} -> IO.puts("Error getting grid lines: #{reason}")
    end

    # Check selection
    case UmyaSpreadsheet.get_selection(spreadsheet, sheet_name) do
      {:ok, selection} ->
        IO.puts("Active Cell: #{selection["active_cell"]}")
        IO.puts("Selected Range: #{selection["sqref"]}")
      {:error, reason} -> IO.puts("Error getting selection: #{reason}")
    end
  end
end

# Usage
SheetInspector.inspect_sheet_settings(spreadsheet, "Sheet1")
```

#### Copying Settings Between Sheets

```elixir
defmodule SheetSettingsCopier do
  def copy_view_settings(spreadsheet, from_sheet, to_sheet) do
    # Copy tab color
    case UmyaSpreadsheet.get_tab_color(spreadsheet, from_sheet) do
      {:ok, color} when color != "" ->
        UmyaSpreadsheet.set_tab_color(spreadsheet, to_sheet, color)
      _ -> :ok
    end

    # Copy view type
    case UmyaSpreadsheet.get_sheet_view(spreadsheet, from_sheet) do
      {:ok, view_type} ->
        UmyaSpreadsheet.set_sheet_view(spreadsheet, to_sheet, view_type)
      _ -> :ok
    end

    # Copy zoom level
    case UmyaSpreadsheet.get_zoom_scale(spreadsheet, from_sheet) do
      {:ok, zoom} ->
        UmyaSpreadsheet.set_zoom_scale(spreadsheet, to_sheet, zoom)
      _ -> :ok
    end

    # Copy grid lines setting
    case UmyaSpreadsheet.get_show_grid_lines(spreadsheet, from_sheet) do
      {:ok, show_grid} ->
        UmyaSpreadsheet.set_show_grid_lines(spreadsheet, to_sheet, show_grid)
      _ -> :ok
    end

    IO.puts("View settings copied from #{from_sheet} to #{to_sheet}")
  end
end

# Usage
SheetSettingsCopier.copy_view_settings(spreadsheet, "Template", "NewSheet")
```

#### Validating Sheet Settings

```elixir
defmodule SheetValidator do
  def validate_standard_settings(spreadsheet, sheet_name) do
    errors = []

    # Check if zoom is within reasonable range
    case UmyaSpreadsheet.get_zoom_scale(spreadsheet, sheet_name) do
      {:ok, zoom} when zoom < 50 or zoom > 200 ->
        errors = ["Zoom level #{zoom}% may be too extreme" | errors]
      _ -> :ok
    end

    # Warn about unusual view types for data sheets
    case UmyaSpreadsheet.get_sheet_view(spreadsheet, sheet_name) do
      {:ok, "page_break_preview"} ->
        errors = ["Page break preview may not be ideal for data entry" | errors]
      _ -> :ok
    end

    if length(errors) > 0 do
      IO.puts("Validation warnings for #{sheet_name}:")
      Enum.each(errors, &IO.puts("  - #{&1}"))
    else
      IO.puts("Sheet #{sheet_name} has standard settings")
    end
  end
end

# Usage
SheetValidator.validate_standard_settings(spreadsheet, "Sheet1")
```

### Default Values

When sheets are created, the following default values are applied:

- **Tab Color**: No color (empty string)
- **Sheet View**: "normal"
- **Zoom Scale**: 100%
- **Grid Lines**: Visible (true)
- **Selection**: Cell A1 with range "A1"

### Performance Notes

- Getter functions are lightweight and can be called frequently
- No performance impact from checking current settings before applying changes
- All getter functions return immediately with cached values
