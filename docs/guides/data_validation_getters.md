# Data Validation Getters Guide

This guide explains how to use the data validation getter functions in UmyaSpreadsheet to retrieve and inspect data validation rules in your Excel files.

## Introduction

Data validation rules help control what data can be entered into cells. UmyaSpreadsheet provides functions to retrieve these rules, allowing you to:

- Check what validation rules exist in a spreadsheet
- Extract specific types of rules (list, number, date, etc.)
- Inspect validation properties (ranges, constraints, messages, etc.)

All getter functions return values in the format `{:ok, result}`, where `result` is the requested information.

## Getting All Data Validation Rules

To retrieve all data validation rules in a sheet:

```elixir
alias UmyaSpreadsheet.{DataValidation}

# Create or open a spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new()
# OR
{:ok, spreadsheet} = UmyaSpreadsheet.read("my_spreadsheet.xlsx")

# Get all data validations
{:ok, all_rules} = DataValidation.get_data_validations(spreadsheet, "Sheet1")
```

The resulting list will contain maps, each representing a validation rule with properties like:

- `range`: The cell range the validation applies to
- `rule_type`: The type of validation (`:list`, `:decimal`, `:date`, etc.)
- Other properties specific to the validation type

## Getting Specific Types of Validation Rules

UmyaSpreadsheet provides specialized getter functions for each validation type:

### List Validations (Dropdowns)

```elixir
# Get all dropdown validations
{:ok, list_rules} = DataValidation.get_list_validations(spreadsheet, "Sheet1")

# Each rule includes the list items
list_items = hd(list_rules).list_items  # ["Option1", "Option2", ...]
```

### Number Validations

```elixir
# Get all number validations
{:ok, number_rules} = DataValidation.get_number_validations(spreadsheet, "Sheet1")

# Example of inspecting a number validation rule
if length(number_rules) > 0 do
  rule = List.first(number_rules)
  min_value = rule.value1  # The minimum value (for "between" validations)
  max_value = rule.value2  # The maximum value (for "between" validations)
  operator = rule.operator # The comparison operator (:between, :greater_than, etc.)
end

# Example printing the constraints
if length(number_rules) > 0 do
  rule = hd(number_rules)
  IO.puts("Number must be #{rule.operator} #{rule.value1} and #{rule.value2}")
end
```

### Date Validations

```elixir
# Get all date validations
{:ok, date_rules} = DataValidation.get_date_validations(spreadsheet, "Sheet1")

# Each rule includes the operator and date values
if length(date_rules) > 0 do
  rule = hd(date_rules)
  IO.puts("Date must be #{rule.operator} #{rule.date1} and #{rule.date2}")
end
```

### Text Length Validations

```elixir
# Get all text length validations
{:ok, length_rules} = DataValidation.get_text_length_validations(spreadsheet, "Sheet1")

# Each rule includes the operator and length values
if length(length_rules) > 0 do
  rule = hd(length_rules)
  IO.puts("Text length must be #{rule.operator} #{rule.length1}")
end
```

### Custom Formula Validations

```elixir
# Get all custom formula validations
{:ok, formula_rules} = DataValidation.get_custom_validations(spreadsheet, "Sheet1")

# Each rule includes the custom formula
if length(formula_rules) > 0 do
  formula = hd(formula_rules).formula
  IO.puts("Custom validation formula: #{formula}")
end
```

## Filtering Validation Rules by Range

Each getter function accepts an optional cell range parameter that filters the results:

```elixir
# Get only validation rules that apply to range A1:B10
{:ok, rules} = DataValidation.get_data_validations(spreadsheet, "Sheet1", "A1:B10")
```

## Utility Functions

### Checking If Validations Exist

```elixir
# Check if any validation rules exist
{:ok, has_rules} = DataValidation.has_data_validations(spreadsheet, "Sheet1")
if has_rules do
  IO.puts("Sheet has validation rules")
end

# Check if a specific range has validation rules
{:ok, has_rules_in_range} = DataValidation.has_data_validations(spreadsheet, "Sheet1", "A1:A10")
```

### Counting Validation Rules

```elixir
# Count all validation rules in a sheet
{:ok, count} = DataValidation.count_data_validations(spreadsheet, "Sheet1")
IO.puts("Sheet has #{count} validation rules")

# Count rules for a specific range
{:ok, range_count} = DataValidation.count_data_validations(spreadsheet, "Sheet1", "A1:A10")
```

## Common Properties

All validation rules include these common properties:

- `range`: The cell range the validation applies to
- `rule_type`: The type of validation
- `allow_blank`: Whether blank values are allowed
- `show_error_message`: Whether error messages are shown
- `error_title`: The title of the error message (if any)
- `error_message`: The error message text (if any)
- `show_input_message`: Whether input prompts are shown
- `prompt_title`: The title of the input prompt (if any)
- `prompt_message`: The input prompt text (if any)

## Example: Analyzing All Validation Types in a Sheet

```elixir
alias UmyaSpreadsheet.DataValidation

{:ok, spreadsheet} = UmyaSpreadsheet.read("my_spreadsheet.xlsx")

# Print a summary of all validation rules
sheet_name = "Sheet1"
{:ok, rules} = DataValidation.get_data_validations(spreadsheet, sheet_name)

IO.puts("Found #{length(rules)} validation rules in #{sheet_name}")

Enum.each(rules, fn rule ->
  IO.puts("\nValidation rule for range: #{rule.range}")
  IO.puts("Type: #{rule.rule_type}")

  case rule.rule_type do
    :list ->
      IO.puts("List items: #{Enum.join(rule.list_items, ", ")}")

    type when type in [:decimal, :whole] ->
      IO.puts("Number must be #{rule.operator} #{rule.value1}#{if rule.value2, do: " and #{rule.value2}", else: ""}")

    :date ->
      IO.puts("Date must be #{rule.operator} #{rule.date1}#{if rule.date2, do: " and #{rule.date2}", else: ""}")

    :text_length ->
      IO.puts("Text length must be #{rule.operator} #{rule.length1}#{if rule.length2, do: " and #{rule.length2}", else: ""}")

    :custom ->
      IO.puts("Custom formula: #{rule.formula}")

    _ ->
      IO.puts("Other validation type")
  end

  if rule.show_error_message do
    IO.puts("Error message: #{rule.error_title} - #{rule.error_message}")
  end

  if rule.show_input_message do
    IO.puts("Input prompt: #{rule.prompt_title} - #{rule.prompt_message}")
  end
end)
```

This guide covers all you need to know about retrieving and inspecting data validation rules in your Excel spreadsheets using UmyaSpreadsheet.
