# Conditional Formatting

This guide covers the conditional formatting capabilities of UmyaSpreadsheet, which allow you to apply formatting rules based on cell values or content.

## Overview

Conditional formatting is a powerful way to visualize data by automatically formatting cells that meet specific conditions. UmyaSpreadsheet supports several types of conditional formatting:

- **Cell Value Rules**: Format cells based on comparison with values (greater than, less than, equal to, etc.)
- **Color Scales**: Apply color gradients based on cell values
- **Data Bars**: Show horizontal bars in cells proportional to their values
- **Icon Sets**: Display different icons based on cell values for visual data representation
- **Top/Bottom Rules**: Highlight top or bottom values/percentages in a range
- **Above/Below Average Rules**: Highlight cells above or below the average value of a range
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

## Icon Sets

Icon sets display different icons in cells based on their values, providing a visual cue for data analysis.

```elixir
# Basic icon set (3 icons)
ConditionalFormatting.add_icon_set(
  spreadsheet,
  "Sheet1",
  "A1:A10",
  "3Symbols",  # Icon set type
  nil,          # No specific value range
  nil           # Default icon style
)

# Custom icon set with specific values
ConditionalFormatting.add_icon_set(
  spreadsheet,
  "Sheet1",
  "A1:A10",
  "3TrafficLights",  # Icon set type
  {"number", "0"},   # Minimum value for red light
  {"number", "100"}, # Maximum value for green light
  nil                 # Default icon style
)
```

### Available Icon Set Types

- `"3Symbols"` - Three symbols (e.g., up/down arrows)
- `"3TrafficLights"` - Traffic light icons (red, yellow, green)
- `"4Ratings"` - Four rating icons (e.g., star ratings)

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

## Above/Below Average Rules

Above/below average rules highlight cells that are above or below the average value of a range, helping to identify outliers.

```elixir
# Highlight cells above average in green
ConditionalFormatting.add_above_below_average_rule(
  spreadsheet,
  "Sheet1",
  "A1:A10",
  "aboveAverage",  # Rule type
  nil,              # No specific value
  "#00FF00"        # Green highlight
)

# Highlight cells below average in red
ConditionalFormatting.add_above_below_average_rule(
  spreadsheet,
  "Sheet1",
  "A1:A10",
  "belowAverage",  # Rule type
  nil,              # No specific value
  "#FF0000"        # Red highlight
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

## Getting Conditional Formatting Rules

UmyaSpreadsheet provides a set of functions to retrieve conditional formatting rules from a spreadsheet. These getter functions allow you to examine the rules applied to specific sheets or ranges.

```elixir
alias UmyaSpreadsheet.ConditionalFormatting

# Get all conditional formatting rules in a sheet
rules = ConditionalFormatting.get_conditional_formatting_rules(spreadsheet, "Sheet1")

# Get conditional formatting rules for a specific range
rules = ConditionalFormatting.get_conditional_formatting_rules(spreadsheet, "Sheet1", "A1:A10")

# Get rules of specific types - these functions return {:ok, list} tuples
cell_value_rules = ConditionalFormatting.get_cell_value_rules(spreadsheet, "Sheet1")
{:ok, color_scales} = ConditionalFormatting.get_color_scales(spreadsheet, "Sheet1")
{:ok, data_bars} = ConditionalFormatting.get_data_bars(spreadsheet, "Sheet1")
{:ok, icon_sets} = ConditionalFormatting.get_icon_sets(spreadsheet, "Sheet1")
top_bottom_rules = ConditionalFormatting.get_top_bottom_rules(spreadsheet, "Sheet1")
above_below_average_rules = ConditionalFormatting.get_above_below_average_rules(spreadsheet, "Sheet1")
text_rules = ConditionalFormatting.get_text_rules(spreadsheet, "Sheet1")
```

### Getter Response Schemas

**Return Type Consistency**: As of version 0.6.17, the specific type getter functions (`get_color_scales`, `get_data_bars`, and `get_icon_sets`) now consistently return `{:ok, list()}` tuples for successful results, maintaining consistency with other getter patterns in the library.

Each of the getter functions returns either a list of maps directly (for general getters) or `{:ok, list}` tuples (for specific type getters), with each map representing a conditional formatting rule. The structure of these maps depends on the type of rule. Below are the detailed schemas for each rule type.

#### Cell Value Rules

`get_cell_value_rules/2` and `get_cell_value_rules/3` return a list of maps with the following structure:

```elixir
%{
  range: String.t(),           # The cell range to which the rule applies, e.g., "A1:A10"
  rule_type: :cell_is,         # Always :cell_is for cell value rules
  operator: atom(),            # One of: :equal, :not_equal, :greater_than, :less_than,
                              # :greater_than_or_equal, :less_than_or_equal, :between, :not_between
  formula: String.t(),         # The formula or value to compare against
  format_style: String.t()     # The color applied as formatting (ARGB format, e.g., "FFFF0000" for red)
}
```

Example:

```elixir
%{
  range: "A1:A10",
  rule_type: :cell_is,
  operator: :greater_than,
  formula: "50",
  format_style: "FFFF0000"
}
```

#### Color Scale Rules

`get_color_scales/2` and `get_color_scales/3` return `{:ok, list}` tuples where each map in the list has the following structure:

```elixir
%{
  range: String.t(),             # The cell range to which the rule applies
  rule_type: :color_scale,       # Always :color_scale for color scale rules
  min_type: atom(),              # The type of the minimum value
  min_value: String.t(),         # The minimum value (if applicable)
  min_color: %{argb: String.t()}, # The color for minimum values (ARGB format)

  # For three-color scales only:
  mid_type: atom(),              # The type of the mid-point value
  mid_value: String.t(),         # The mid-point value (if applicable)
  mid_color: %{argb: String.t()}, # The color for mid-point values (ARGB format)

  max_type: atom(),              # The type of the maximum value
  max_value: String.t(),         # The maximum value (if applicable)
  max_color: %{argb: String.t()}  # The color for maximum values (ARGB format)
}
```

Value type atoms can be: `:min`, `:max`, `:number`, `:percent`, `:percentile`, or `:formula`

Example (two-color scale):

```elixir
%{
  range: "A1:A10",
  rule_type: :color_scale,
  min_type: :min,
  min_value: "",
  min_color: %{argb: "FF0000FF"},  # Blue
  max_type: :max,
  max_value: "",
  max_color: %{argb: "FF00FF00"}   # Green
}
```

Example (three-color scale):

```elixir
%{
  range: "A1:A10",
  rule_type: :color_scale,
  min_type: :min,
  min_value: "",
  min_color: %{argb: "FFFF0000"},  # Red
  mid_type: :percentile,
  mid_value: "50",
  mid_color: %{argb: "FFFFFF00"},  # Yellow
  max_type: :max,
  max_value: "",
  max_color: %{argb: "FF00FF00"}   # Green
}
```

#### Data Bar Rules

`get_data_bars/2` and `get_data_bars/3` return `{:ok, list}` tuples where each map in the list has the following structure:

```elixir
%{
  range: String.t(),    # The cell range to which the rule applies
  rule_type: :data_bar, # Always :data_bar for data bar rules
  min_value: term(),    # The minimum value as a tuple like {type, value} or nil
  max_value: term(),    # The maximum value as a tuple like {type, value} or nil
  color: String.t()     # The color of the data bar (ARGB format)
}
```

The `min_value` and `max_value` can be:

- A tuple like `{:min, ""}` or `{:number, "20"}`
- `nil` when using default values

Example:

```elixir
%{
  range: "C1:C10",
  rule_type: :data_bar,
  min_value: nil,  # Using default min (lowest value in range)
  max_value: nil,  # Using default max (highest value in range)
  color: "FF638EC6"  # Blue data bar
}
```

#### Icon Set Rules

`get_icon_sets/2` and `get_icon_sets/3` return `{:ok, list}` tuples where each map in the list has the following structure:

```elixir
%{
  range: String.t(),       # The cell range to which the rule applies
  rule_type: :icon_set,    # Always :icon_set for icon set rules
  icon_style: String.t(),  # The icon set type, e.g., "3arrows", "4arrows", "5arrows"
  thresholds: [{String.t(), String.t()}]  # List of threshold values as {type, value} tuples
}
```

Example:

```elixir
%{
  range: "A1:A10",
  rule_type: :icon_set,
  icon_style: "3arrows",
  thresholds: [
    {"percent", "0"},
    {"percent", "33"},
    {"percent", "67"}
  ]
}
```

#### Top/Bottom Rules

`get_top_bottom_rules/2` and `get_top_bottom_rules/3` return a list of maps with the following structure:

```elixir
%{
  range: String.t(),          # The cell range to which the rule applies
  rule_type: :top_bottom,     # Always :top_bottom for top/bottom rules
  rule_type_value: atom(),    # Either :top or :bottom
  rank: integer(),            # The number of values or percentage to highlight
  percent: boolean(),         # Whether rank is a percentage or count
  format_style: String.t()    # The color applied as formatting (ARGB format)
}
```

Example:

```elixir
%{
  range: "A1:A10",
  rule_type: :top_bottom,
  rule_type_value: :top,
  rank: 3,
  percent: false,  # Top 3 values, not percentile
  format_style: "FFFFFF00"  # Yellow highlight
}
```

#### Above/Below Average Rules

`get_above_below_average_rules/2` and `get_above_below_average_rules/3` return a list of maps with the following structure:

```elixir
%{
  range: String.t(),              # The cell range to which the rule applies
  rule_type: :above_below_average, # Always :above_below_average for these rules
  rule_type_value: atom(),        # One of: :above, :below, :above_equal, :below_equal
  std_dev: integer(),             # Standard deviation value (0 for regular above/below)
  format_style: String.t()        # The color applied as formatting (ARGB format)
}
```

Example:

```elixir
%{
  range: "C1:C10",
  rule_type: :above_below_average,
  rule_type_value: :above,
  std_dev: 0,  # 0 means standard above average, not standard deviation based
  format_style: "FF00FF00"  # Green highlight
}
```

#### Text Rules

`get_text_rules/2` and `get_text_rules/3` return a list of maps with the following structure:

```elixir
%{
  range: String.t(),      # The cell range to which the rule applies
  rule_type: :text_rule,  # Always :text_rule for text rules
  operator: atom(),       # One of: :contains, :not_contains, :begins_with, :ends_with
  text: String.t(),       # The text to search for
  format_style: String.t() # The color applied as formatting (ARGB format)
}
```

Example:

```elixir
%{
  range: "B1:B10",
  rule_type: :text_rule,
  operator: :begins_with,
  text: "Value 1",
  format_style: "FFFFC000"  # Orange highlight
}
```

#### All Rules

When using `get_conditional_formatting_rules/2` or `get_conditional_formatting_rules/3`, you'll get a list containing any of the above rule types, each with their respective structure. You can identify the rule type by checking the `:rule_type` field in each map.

### Working with Rule Results

The getter functions return a list of rule maps, each containing information about a conditional formatting rule:

```elixir
# Example: Getting cell value rules
cell_value_rules = ConditionalFormatting.get_cell_value_rules(spreadsheet, "Sheet1")

# Each rule contains details about the formatting condition
Enum.each(cell_value_rules, fn rule ->
  IO.puts("Range: #{rule.range}")
  IO.puts("Operator: #{rule.operator}")
  IO.puts("Formula: #{rule.formula}")
  IO.puts("Format style: #{rule.format_style}")
end)
```

### Example: Analyzing Conditional Formatting Rules

```elixir
alias UmyaSpreadsheet.ConditionalFormatting

# Read a spreadsheet with conditional formatting
{:ok, spreadsheet} = UmyaSpreadsheet.read("formatted_data.xlsx")

# Get all conditional formatting rules
all_rules = ConditionalFormatting.get_conditional_formatting_rules(spreadsheet, "Sheet1")

# Count rules by type
rule_counts = Enum.reduce(all_rules, %{}, fn rule, acc ->
  Map.update(acc, rule.type, 1, &(&1 + 1))
end)

IO.puts("Rule counts by type:")
Enum.each(rule_counts, fn {type, count} ->
  IO.puts("#{type}: #{count}")
end)

# Find all red formatting
red_rules = Enum.filter(all_rules, fn rule ->
  rule.format["bgColor"] in ["#FF0000", "rgb(255,0,0)"]
end)

IO.puts("\nRed formatting rules:")
Enum.each(red_rules, fn rule ->
  IO.puts("#{rule.type} rule on range #{rule.range}")
end)
```

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

# Read back and analyze the rules
{:ok, read_spreadsheet} = UmyaSpreadsheet.read("conditional_formatting_example.xlsx")
rules = ConditionalFormatting.get_conditional_formatting_rules(read_spreadsheet, "Sheet1")

IO.puts("Found #{length(rules)} conditional formatting rules in Sheet1")
```

## Notes

- Conditional formatting is applied to a range of cells, not just individual cells
- Multiple rules can be applied to the same range, and they will be evaluated in the order they were added
- The formatting is applied when opening the file in Excel, not in the Elixir code itself
- For complex formatting needs, consider using multiple rules to achieve the desired effect
- Use the getter functions to inspect and analyze conditional formatting rules in existing spreadsheets
