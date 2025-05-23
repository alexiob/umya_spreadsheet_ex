# Conditional Formatting

This guide covers the conditional formatting capabilities of UmyaSpreadsheet, which allow you to apply formatting rules based on cell values or content.

## Overview

Conditional formatting is a powerful way to visualize data by automatically formatting cells that meet specific conditions. UmyaSpreadsheet supports several types of conditional formatting:

- **Cell Value Rules**: Format cells based on comparison with values (greater than, less than, equal to, etc.)
- **Color Scales**: Apply color gradients based on cell values
- **Data Bars**: Show horizontal bars in cells proportional to their values
- **Top/Bottom Rules**: Highlight top or bottom values/percentages in a range
- **Text Rules**: Format cells based on text content (contains, begins with, etc.)

All conditional formatting functionality is available through the `UmyaSpreadsheet.ConditionalFormatting` module.

## Cell Value Rules

Cell value rules allow you to apply formatting based on cell values that meet specific criteria.

```elixir
alias UmyaSpreadsheet.ConditionalFormatting

# Highlight cells with values greater than 50 in red
ConditionalFormatting.add_cell_value_rule(
  spreadsheet,
  "Sheet1",
  "A1:A10",  # Range to apply formatting to
  "greaterThan",  # Operator
  "50",  # Value to compare against
  nil,  # Second value (only used for "between" and "notBetween")
  "#FF0000"  # Format to apply (red background)
)

# Highlight cells with values between 20 and 80 in green
ConditionalFormatting.add_cell_value_rule(
  spreadsheet,
  "Sheet1",
  "A1:A10",
  "between",
  "20",  # First value
  "80",  # Second value
  "#00FF00"  # Format to apply (green background)
)
```

### Available Operators

- `"equal"` - Equal to value
- `"notEqual"` - Not equal to value
- `"greaterThan"` - Greater than value
- `"greaterThanOrEqual"` - Greater than or equal to value
- `"lessThan"` - Less than value
- `"lessThanOrEqual"` - Less than or equal to value
- `"between"` - Between two values (requires second value)
- `"notBetween"` - Not between two values (requires second value)

## Color Scales

Color scales create a gradient of colors based on cell values, making it easy to visualize the distribution of values in a range.

```elixir
# Two-color scale (red to green)
ConditionalFormatting.add_color_scale(
  spreadsheet,
  "Sheet1",
  "A1:A10",
  nil,  # Use default min (minimum value in range)
  nil,  # No mid point
  nil,  # Use default max (maximum value in range)
  "#FF0000",  # Red for minimum
  nil,        # No mid color
  "#00FF00"   # Green for maximum
)

# Three-color scale (red to yellow to green)
ConditionalFormatting.add_color_scale(
  spreadsheet,
  "Sheet1",
  "A1:A10",
  {"min", ""},         # Use minimum value
  {"percentile", "50"}, # Use 50th percentile for mid-point
  {"max", ""},         # Use maximum value
  "#FF0000",           # Red for minimum
  "#FFFF00",           # Yellow for middle
  "#00FF00"            # Green for maximum
)
```

### Value Types

For min, mid, and max values, you can use:

- `{"min", ""}` - Use the minimum value in the range
- `{"max", ""}` - Use the maximum value in the range
- `{"percentile", "X"}` - Use the Xth percentile (e.g., "50" for median)
- `{"percent", "X"}` - Use X percent of the range
- `{"number", "X"}` - Use a specific number
- `{"formula", "X"}` - Use a formula result

## Data Bars

Data bars create horizontal bars in each cell with length proportional to the cell value, providing a visual representation of the data.

```elixir
# Basic data bars with default min/max values
ConditionalFormatting.add_data_bar(
  spreadsheet,
  "Sheet1",
  "A1:A10",
  nil,  # Use default min (auto)
  nil,  # Use default max (auto)
  "#6C8EBF"  # Blue bars
)

# Data bars with custom min/max values
ConditionalFormatting.add_data_bar(
  spreadsheet,
  "Sheet1",
  "A1:A10",
  {"number", "20"},  # Minimum value is 20
  {"number", "80"},  # Maximum value is 80
  "#FF9900"          # Orange bars
)
```

## Top/Bottom Rules

Top/bottom rules highlight cells with values that rank at the top or bottom of a range, either by count or percentage.

```elixir
# Highlight top 3 values in red
ConditionalFormatting.add_top_bottom_rule(
  spreadsheet,
  "Sheet1",
  "A1:A10",
  "top",    # Rule type
  3,        # Top 3 values
  false,    # Not a percentage
  "#FF0000" # Red highlight
)

# Highlight bottom 20% in yellow
ConditionalFormatting.add_top_bottom_rule(
  spreadsheet,
  "Sheet1",
  "A1:A10",
  "bottom", # Rule type
  20,       # Bottom 20%
  true,     # Is a percentage
  "#FFFF00" # Yellow highlight
)
```

## Text Rules

Text rules allow you to apply formatting based on the text content of cells.

```elixir
# Highlight cells containing "Error" in red
ConditionalFormatting.add_text_rule(
  spreadsheet,
  "Sheet1",
  "A1:A10",
  "contains",  # Rule type
  "Error",     # Text to search for
  "#FF0000"    # Red highlight
)

# Highlight cells beginning with "Warning" in yellow
ConditionalFormatting.add_text_rule(
  spreadsheet,
  "Sheet1",
  "A1:A10",
  "beginsWith", # Rule type
  "Warning",    # Text to search for
  "#FFFF00"     # Yellow highlight
)
```

### Available Text Rule Types

- `"contains"` - Cell contains the text
- `"notContains"` - Cell does not contain the text
- `"beginsWith"` - Cell begins with the text
- `"endsWith"` - Cell ends with the text

## Format Styles

For all conditional formatting rules, you can specify a format to apply when the condition is met. Currently, the implementation supports specifying a background color in hexadecimal format or RGB format:

- Hexadecimal: `"#FF0000"` (red), `"#00FF00"` (green), `"#0000FF"` (blue)
- RGB: `"rgb(255,0,0)"` (red), `"rgb(0,255,0)"` (green), `"rgb(0,0,255)"` (blue)

## Complete Example

Here's a comprehensive example that applies multiple conditional formatting rules to a worksheet:

```elixir
alias UmyaSpreadsheet.ConditionalFormatting

# Create a new spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add test data
for i <- 1..10 do
  UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A#{i}", "#{i * 10}")
end

# Cell value rule: values less than 30 (red)
ConditionalFormatting.add_cell_value_rule(
  spreadsheet, "Sheet1", "A1:A10", "lessThan", "30", nil, "#FF0000"
)

# Cell value rule: values between 30 and 70 (yellow)
ConditionalFormatting.add_cell_value_rule(
  spreadsheet, "Sheet1", "A1:A10", "between", "30", "70", "#FFFF00"
)

# Cell value rule: values greater than 70 (green)
ConditionalFormatting.add_cell_value_rule(
  spreadsheet, "Sheet1", "A1:A10", "greaterThan", "70", nil, "#00FF00"
)

# Add a new sheet for color scales
UmyaSpreadsheet.add_sheet(spreadsheet, "ColorScales")
for i <- 1..10 do
  UmyaSpreadsheet.set_cell_value(spreadsheet, "ColorScales", "A#{i}", "#{i * 10}")
end

# Add a three-color scale
ConditionalFormatting.add_color_scale(
  spreadsheet, "ColorScales", "A1:A10",
  {"min", ""}, {"percentile", "50"}, {"max", ""},
  "#FF0000", "#FFFF00", "#00FF00"
)

# Save the spreadsheet
UmyaSpreadsheet.write(spreadsheet, "conditional_formatting_example.xlsx")
```

## Notes

- Conditional formatting is applied to a range of cells, not just individual cells
- Multiple rules can be applied to the same range, and they will be evaluated in the order they were added
- The formatting is applied when opening the file in Excel, not in the Elixir code itself
- For complex formatting needs, consider using multiple rules to achieve the desired effect
