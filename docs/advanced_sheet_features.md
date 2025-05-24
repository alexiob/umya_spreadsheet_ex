# Advanced Sheet Features

This guide explores the advanced sheet-level features available in UmyaSpreadsheet for customizing how Excel sheets are displayed and interacted with.

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
