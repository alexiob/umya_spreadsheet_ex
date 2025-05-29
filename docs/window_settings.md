# Window Settings

This guide covers the window settings functionality in UmyaSpreadsheet, allowing you to control how an Excel workbook and its worksheets appear when opened in Excel.

## Sheet Selection

You can set which cell or range is selected when a user opens a particular worksheet:

```elixir
alias UmyaSpreadsheet

# Create a new spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add some data
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Header")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", "Content")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C3", "More Content")

# Set the active cell to B2 when the sheet is opened
UmyaSpreadsheet.set_selection(spreadsheet, "Sheet1", "B2")

# Or select a range of cells (the first cell becomes the active cell)
UmyaSpreadsheet.set_selection(spreadsheet, "Sheet1", "B2:C5")

# Save the spreadsheet
UmyaSpreadsheet.write(spreadsheet, "window_settings.xlsx")
```

## Working with Active Tabs

When a workbook contains multiple worksheets, you can control which sheet is active when the workbook is opened:

```elixir
# Create a new sheet
UmyaSpreadsheet.add_sheet(spreadsheet, "Data")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Data", "A1", "This is the data sheet")

# Create another sheet
UmyaSpreadsheet.add_sheet(spreadsheet, "Summary")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Summary", "A1", "This is the summary sheet")

# Set "Summary" as the active sheet when the workbook is opened
# Sheet indices are 0-based, so if "Summary" is the third sheet, its index is 2
UmyaSpreadsheet.set_active_tab(spreadsheet, 2)
```

To determine the index of a sheet, you can use:

```elixir
sheet_names = UmyaSpreadsheet.get_sheet_names(spreadsheet)
```

You can also retrieve the current active tab index using:

```elixir
# Get the active tab index (returns an integer, 0-based)
active_tab_index = UmyaSpreadsheet.get_active_tab(spreadsheet)

# Use this to determine which sheet is currently active
sheet_names = UmyaSpreadsheet.get_sheet_names(spreadsheet)
active_sheet_name = Enum.at(sheet_names, active_tab_index)

IO.puts("Active sheet: #{active_sheet_name}")
```

## Window Position and Size

You can control the initial position and size of the Excel window when a workbook is opened:

```elixir
# Set the window position (x, y) and size (width, height)
UmyaSpreadsheet.set_workbook_window_position(spreadsheet, 100, 50, 800, 600)
```

To retrieve the current window position and size settings:

```elixir
# Get the window position and size as a map
position = UmyaSpreadsheet.get_workbook_window_position(spreadsheet)

# The returned map has :x, :y, :width, and :height keys
IO.puts("Window position: #{position[:x]}, #{position[:y]}")
IO.puts("Window size: #{position[:width]} x #{position[:height]}")
```

summary_index = Enum.find_index(sheet_names, fn name -> name == "Summary" end)
UmyaSpreadsheet.set_active_tab(spreadsheet, summary_index)

```

## Workbook Window Position and Size

You can control the position and size of the Excel application window when the workbook is opened:

```elixir
# Set the window position and size (all values in pixels)
# Parameters: left, top, width, height
UmyaSpreadsheet.set_workbook_window_position(spreadsheet, 100, 100, 800, 600)
```

This setting only works when:

- The workbook is opened in Excel desktop applications
- The user does not have Excel already open
- The user's Excel preferences don't override these settings

## Combining Window Settings

You can combine multiple window settings to create a specific user experience:

```elixir
# Create a new spreadsheet with multiple sheets
{:ok, spreadsheet} = UmyaSpreadsheet.new()
UmyaSpreadsheet.add_sheet(spreadsheet, "Data")
UmyaSpreadsheet.add_sheet(spreadsheet, "Analysis")

# Add data to sheets
# ...

# Configure window settings
UmyaSpreadsheet.set_active_tab(spreadsheet, 1)  # "Analysis" sheet is active
UmyaSpreadsheet.set_selection(spreadsheet, "Analysis", "B5")  # B5 is selected
UmyaSpreadsheet.set_workbook_window_position(spreadsheet, 50, 50, 1200, 800)  # Large window

# Save the spreadsheet
UmyaSpreadsheet.write(spreadsheet, "configured_workbook.xlsx")
```

## Best Practices

1. **User-Friendly Defaults**: Set the active tab to the most relevant sheet for users
2. **Focus Attention**: Use selection to direct users to important cells or data entry areas
3. **Screen Compatibility**: When setting window positions and sizes, consider varied screen resolutions
4. **Testing**: Test your settings on different Excel versions to ensure compatibility
5. **Documentation**: Document any specific window settings in your code for future reference

## Limitations

- Some Excel settings may override workbook window settings
- Window position settings may behave differently across Excel versions and platforms
- Window settings are suggestions to Excel and may not be honored in all circumstances
