# Sheet View Settings

This guide covers the sheet view functionality in UmyaSpreadsheet, allowing you to control how individual worksheets appear and behave when opened in Excel, including zoom levels, grid lines, pane freezing, and tab colors.

## Overview

Sheet view settings control the visual appearance and user interaction experience for individual worksheets. These settings are applied per sheet and include:

- **Zoom levels** for different view modes
- **Grid line visibility** controls
- **Pane freezing and splitting** for data navigation
- **Tab colors** for visual organization
- **Cell selection** management

## Zoom Settings

Control the zoom level for different Excel view modes. Excel supports different zoom levels for Normal view, Page Layout view, and Page Break Preview mode.

### Setting Zoom Levels

```elixir
alias UmyaSpreadsheet

# Create a new spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Set zoom for normal view (default mode)
UmyaSpreadsheet.set_zoom_scale_normal(spreadsheet, "Sheet1", 125)

# Set zoom for page layout view
UmyaSpreadsheet.set_zoom_scale_page_layout(spreadsheet, "Sheet1", 75)

# Set zoom for page break preview mode
UmyaSpreadsheet.set_zoom_scale_page_break(spreadsheet, "Sheet1", 85)

# Save the spreadsheet
UmyaSpreadsheet.write(spreadsheet, "zoom_settings.xlsx")
```

### Getting Current Zoom Levels

```elixir
# Get the current zoom level for normal view
normal_zoom = UmyaSpreadsheet.get_zoom_scale(spreadsheet, "Sheet1")
IO.puts("Normal view zoom: #{normal_zoom}%")  # Default: 100%

# Get zoom for page layout view
page_layout_zoom = UmyaSpreadsheet.get_zoom_scale_page_layout(spreadsheet, "Sheet1")

# Get zoom for page break preview
page_break_zoom = UmyaSpreadsheet.get_zoom_scale_page_break(spreadsheet, "Sheet1")
```

**Zoom Level Guidelines:**

- Valid range: 10% to 400%
- Default value: 100%
- Values are automatically clamped to valid range
- Common useful values: 75%, 100%, 125%, 150%

## Grid Lines

Control the visibility of grid lines in worksheets. Grid lines are the light gray lines that separate cells in Excel.

### Managing Grid Line Visibility

```elixir
# Hide grid lines
UmyaSpreadsheet.set_show_grid_lines(spreadsheet, "Sheet1", false)

# Show grid lines (default)
UmyaSpreadsheet.set_show_grid_lines(spreadsheet, "Sheet1", true)

# Check current grid line visibility
grid_lines_visible = UmyaSpreadsheet.get_show_grid_lines(spreadsheet, "Sheet1")
IO.puts("Grid lines visible: #{grid_lines_visible}")  # Default: true
```

**Grid Line Best Practices:**

- Grid lines are shown by default (Excel standard)
- Hide grid lines for cleaner presentation sheets
- Keep grid lines visible for data entry sheets
- Grid lines don't affect printing unless specifically configured

## Pane Freezing and Splitting

Freeze or split panes to keep headers visible while scrolling through large datasets.

### Freezing Panes

```elixir
# Freeze panes at cell C3 (keeps rows 1-2 and columns A-B visible)
UmyaSpreadsheet.freeze_panes(spreadsheet, "Sheet1", "C3", "C3")

# Freeze the top row only
UmyaSpreadsheet.freeze_panes(spreadsheet, "Sheet1", "A2", "A2")

# Freeze the first column only
UmyaSpreadsheet.freeze_panes(spreadsheet, "Sheet1", "B1", "B1")

# Freeze both top row and first column
UmyaSpreadsheet.freeze_panes(spreadsheet, "Sheet1", "B2", "B2")
```

### Splitting Panes

```elixir
# Split panes at specific pixel coordinates
# Parameters: x_split (column position), y_split (row position)
UmyaSpreadsheet.split_panes(spreadsheet, "Sheet1", 2000, 1000)

# Split only vertically (column split)
UmyaSpreadsheet.split_panes(spreadsheet, "Sheet1", 1500, 0)

# Split only horizontally (row split)
UmyaSpreadsheet.split_panes(spreadsheet, "Sheet1", 0, 800)
```

**Freezing vs. Splitting:**

- **Freezing**: Creates fixed regions that don't scroll
- **Splitting**: Creates resizable panes with independent scrolling
- **Use freezing** for headers and labels
- **Use splitting** when you need flexible viewing of different data regions

## Tab Colors

Set colors for worksheet tabs to organize and categorize sheets visually.

### Setting Tab Colors

```elixir
# Set tab color using hex color codes
UmyaSpreadsheet.set_tab_color(spreadsheet, "Sheet1", "#FF0000")  # Red
UmyaSpreadsheet.set_tab_color(spreadsheet, "Summary", "#00FF00")  # Green
UmyaSpreadsheet.set_tab_color(spreadsheet, "Data", "#0000FF")    # Blue

# Set tab color using RGB values
UmyaSpreadsheet.set_tab_color(spreadsheet, "Analysis", "#FFA500")  # Orange

# Remove tab color (reset to default)
UmyaSpreadsheet.set_tab_color(spreadsheet, "Sheet1", "")
```

### Getting Tab Colors

```elixir
# Get the current tab color
tab_color = UmyaSpreadsheet.get_tab_color(spreadsheet, "Sheet1")

case tab_color do
  "" -> IO.puts("No tab color set (default)")
  color -> IO.puts("Tab color: #{color}")
end
```

**Tab Color Tips:**

- Use consistent color schemes across related workbooks
- Common patterns: Red for important/urgent, Green for completed, Blue for data
- Hex format: `#RRGGBB` (e.g., `#FF0000` for red)
- Empty string removes the color
- Colors are preserved when copying/cloning sheets

## Cell Selection

Manage which cells are selected when a sheet is opened.

### Setting Selection

```elixir
# Select a single cell
UmyaSpreadsheet.set_selection(spreadsheet, "Sheet1", "B5")

# Select a range of cells
UmyaSpreadsheet.set_selection(spreadsheet, "Sheet1", "A1:C10")

# Select multiple discontinuous ranges (use semicolon separator)
UmyaSpreadsheet.set_selection(spreadsheet, "Sheet1", "A1:B2;D4:E6")
```

### Getting Current Selection

```elixir
# Get the current selection
selection = UmyaSpreadsheet.get_selection(spreadsheet, "Sheet1")
IO.puts("Selected range: #{selection}")
```

**Selection Best Practices:**

- Direct users to important cells or data entry areas
- Use selection to highlight key information
- For data entry forms, select the first input cell
- Clear and meaningful selections improve user experience

## Complete Example

Here's a comprehensive example that demonstrates multiple sheet view features:

```elixir
alias UmyaSpreadsheet

# Create a new spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add some sample data
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Sales Report")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "Product")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", "Revenue")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C2", "Units Sold")

# Add data rows
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "Widget A")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B3", 5000)
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C3", 50)

# Configure sheet view settings
UmyaSpreadsheet.set_zoom_scale_normal(spreadsheet, "Sheet1", 110)  # Slightly larger
UmyaSpreadsheet.set_show_grid_lines(spreadsheet, "Sheet1", true)   # Keep grid lines
UmyaSpreadsheet.freeze_panes(spreadsheet, "Sheet1", "A3", "A3")    # Freeze headers
UmyaSpreadsheet.set_tab_color(spreadsheet, "Sheet1", "#4CAF50")    # Green tab
UmyaSpreadsheet.set_selection(spreadsheet, "Sheet1", "A3")         # Focus on data

# Create a summary sheet with different settings
UmyaSpreadsheet.add_sheet(spreadsheet, "Summary")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Summary", "A1", "Executive Summary")

# Configure summary sheet view
UmyaSpreadsheet.set_zoom_scale_normal(spreadsheet, "Summary", 125)    # Larger zoom
UmyaSpreadsheet.set_show_grid_lines(spreadsheet, "Summary", false)    # Clean look
UmyaSpreadsheet.set_tab_color(spreadsheet, "Summary", "#2196F3")      # Blue tab
UmyaSpreadsheet.set_selection(spreadsheet, "Summary", "A1")           # Focus on title

# Save the configured spreadsheet
UmyaSpreadsheet.write(spreadsheet, "configured_views.xlsx")

IO.puts("Spreadsheet created with custom view settings!")
```

## Default Values

When creating new spreadsheets or sheets, the following defaults are applied:

- **Zoom Scale**: 100% for all view modes
- **Grid Lines**: Visible (`true`)
- **Tab Color**: None (empty string)
- **Selection**: Cell A1
- **Panes**: Not frozen or split

These defaults match Excel's standard behavior for new worksheets.

## Advanced Usage

### Checking for Existing Settings

```elixir
# Check if a sheet has custom view settings
defmodule SheetViewChecker do
  def has_custom_settings?(spreadsheet, sheet_name) do
    zoom = UmyaSpreadsheet.get_zoom_scale(spreadsheet, sheet_name)
    grid_lines = UmyaSpreadsheet.get_show_grid_lines(spreadsheet, sheet_name)
    tab_color = UmyaSpreadsheet.get_tab_color(spreadsheet, sheet_name)

    zoom != 100 or grid_lines != true or tab_color != ""
  end
end

custom_settings = SheetViewChecker.has_custom_settings?(spreadsheet, "Sheet1")
```

### Copying View Settings Between Sheets

```elixir
defmodule SheetViewCopier do
  def copy_view_settings(spreadsheet, from_sheet, to_sheet) do
    # Copy zoom settings
    zoom = UmyaSpreadsheet.get_zoom_scale(spreadsheet, from_sheet)
    UmyaSpreadsheet.set_zoom_scale_normal(spreadsheet, to_sheet, zoom)

    # Copy grid line setting
    grid_lines = UmyaSpreadsheet.get_show_grid_lines(spreadsheet, from_sheet)
    UmyaSpreadsheet.set_show_grid_lines(spreadsheet, to_sheet, grid_lines)

    # Copy tab color
    tab_color = UmyaSpreadsheet.get_tab_color(spreadsheet, from_sheet)
    UmyaSpreadsheet.set_tab_color(spreadsheet, to_sheet, tab_color)

    IO.puts("View settings copied from #{from_sheet} to #{to_sheet}")
  end
end

SheetViewCopier.copy_view_settings(spreadsheet, "Sheet1", "Sheet2")
```

## Troubleshooting

### Common Issues

1. **Zoom not applying**: Ensure zoom values are between 10 and 400
2. **Grid lines still visible**: Check if you're viewing the correct sheet
3. **Freeze panes not working**: Verify cell references are valid
4. **Tab colors not showing**: Ensure hex codes include the '#' prefix

### Performance Considerations

- Sheet view settings have minimal performance impact
- Settings are applied per sheet, not globally
- Large numbers of frozen panes may affect Excel performance
- Tab colors don't impact file size significantly

## Best Practices

1. **Consistency**: Use consistent view settings across related sheets
2. **User Experience**: Configure settings to highlight important areas
3. **Accessibility**: Avoid color-only communication for tab colors
4. **Performance**: Don't overuse complex pane splitting
5. **Testing**: Test view settings in the target Excel version

## Compatibility

These sheet view features are compatible with:

- Excel 2007 and later (.xlsx format)
- LibreOffice Calc
- Google Sheets (with some limitations)
- Other applications supporting Office Open XML format

Settings are preserved when files are opened in different applications, though exact rendering may vary.
