# Data Validation

This guide explains how to use UmyaSpreadsheet's data validation features to control what users can enter in Excel cells.

## Introduction to Data Validation

Data validation allows you to define rules that restrict what values can be entered in specific cells. This helps ensure data integrity and can provide users with guidance on what they should enter.

### Benefits of Data Validation

- **Data Integrity**: Ensure data entered meets specific requirements
- **User Guidance**: Help users understand what values are expected
- **Error Prevention**: Reduce errors from incorrect data entry
- **Standardization**: Maintain consistency across your spreadsheet

### Types of Validation Supported

UmyaSpreadsheet supports several types of validation:

- **List Validation**: Creates dropdown lists with predefined options
- **Number Validation**: Restricts input to numbers within specific ranges
- **Date Validation**: Restricts input to dates within specific ranges (supports both ISO strings and Date structs)
- **Text Length Validation**: Restricts input to text with specific character count requirements
- **Custom Formula Validation**: Allows advanced validation using Excel formulas

## Adding Data Validation Rules

### List Validation (Dropdown Lists)

Create dropdown lists to restrict input to a predefined set of options:

```elixir
alias UmyaSpreadsheet.{DataValidation}

# Create a dropdown with simple options
DataValidation.add_list_validation(
  spreadsheet,
  "Sheet1",
  "A1:A10",               # Cell range
  ["Red", "Green", "Blue"], # List items
  true,                   # Allow blank (optional)
  "Invalid Selection",    # Error title (optional)
  "Please select a color from the list", # Error message (optional)
  "Color Selection",      # Prompt title (optional)
  "Select a color"        # Prompt message (optional)
)
```

### Number Validation

Restrict input to numbers that meet specific criteria:

```elixir
# Allow only numbers between 1 and 100
DataValidation.add_number_validation(
  spreadsheet,
  "Sheet1",
  "B1:B10",                # Cell range
  "between",               # Operator
  1.0,                     # Minimum value
  100.0,                   # Maximum value
  true,                    # Allow blank (optional)
  "Invalid Number",        # Error title (optional)
  "Please enter a number between 1 and 100", # Error message (optional)
  "Number Input",          # Prompt title (optional)
  "Enter a number between 1 and 100"         # Prompt message (optional)
)

# Allow only values greater than 0
DataValidation.add_number_validation(
  spreadsheet,
  "Sheet1",
  "C1:C10",
  "greaterThan",
  0.0
)
```

### Date Validation

Restrict input to dates that meet specific criteria. You can use either ISO format strings or Elixir Date structs for date parameters:

```elixir
# Allow only dates in the current year (using string dates)
DataValidation.add_date_validation(
  spreadsheet,
  "Sheet1",
  "D1:D10",                # Cell range
  "between",               # Operator
  "#{Date.utc_today().year}-01-01", # Start date as string
  "#{Date.utc_today().year}-12-31", # End date as string
  true,                    # Allow blank (optional)
  "Invalid Date",          # Error title (optional)
  "Please enter a date in #{Date.utc_today().year}", # Error message (optional)
  "Date Input",            # Prompt title (optional)
  "Enter a date in #{Date.utc_today().year}"         # Prompt message (optional)
)

# Allow only future dates (using Date struct directly)
DataValidation.add_date_validation(
  spreadsheet,
  "Sheet1",
  "E1:E10",
  "greaterThanOrEqual",
  Date.utc_today(),  # Date struct works directly
  nil,
  true,
  "Invalid Date",
  "Please enter today's date or later"
)
```

### Text Length Validation

Restrict input to text with specific length requirements:

```elixir
# Limit input to at most 280 characters (like a tweet)
DataValidation.add_text_length_validation(
  spreadsheet,
  "Sheet1",
  "F1:F10",                # Cell range
  "lessThanOrEqual",       # Operator
  280,                     # Maximum length
  nil,                     # Second value (not used)
  true,                    # Allow blank (optional)
  "Text Too Long",         # Error title (optional)
  "Please enter at most 280 characters", # Error message (optional)
  "Character Limit",       # Prompt title (optional)
  "Enter up to 280 characters" # Prompt message (optional)
)

# Require at least 8 characters (like a password)
DataValidation.add_text_length_validation(
  spreadsheet,
  "Sheet1",
  "G1:G10",
  "greaterThanOrEqual",
  8
)
```

### Custom Formula Validation

Create advanced validation rules using Excel formulas:

```elixir
# Only allow values divisible by 5
DataValidation.add_custom_validation(
  spreadsheet,
  "Sheet1",
  "H1:H10",                # Cell range
  "MOD(H1,5)=0",           # Formula (without = sign)
  true,                    # Allow blank (optional)
  "Invalid Value",         # Error title (optional)
  "Please enter a value divisible by 5", # Error message (optional)
  "Value Input",           # Prompt title (optional)
  "Enter a value divisible by 5" # Prompt message (optional)
)
```

## Retrieving and Inspecting Data Validation Rules

UmyaSpreadsheet provides functions to retrieve and inspect data validation rules. All getter functions return values in the format `{:ok, result}`.

### Getting All Data Validation Rules

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

### Getting Specific Types of Validation Rules

UmyaSpreadsheet provides specialized getter functions for each validation type:

#### List Validations (Dropdowns)

```elixir
# Get all dropdown validations
{:ok, list_rules} = DataValidation.get_list_validations(spreadsheet, "Sheet1")

# Each rule includes the list items
list_items = hd(list_rules).list_items  # ["Option1", "Option2", ...]
```

#### Number Validations

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
```

#### Date Validations

```elixir
# Get all date validations
{:ok, date_rules} = DataValidation.get_date_validations(spreadsheet, "Sheet1")

# Each rule includes the operator and date values
if date_rules != [] do
  rule = hd(date_rules)
  IO.puts("Date must be #{rule.operator} #{rule.date1} and #{rule.date2}")
end
```

#### Text Length Validations

```elixir
# Get all text length validations
{:ok, length_rules} = DataValidation.get_text_length_validations(spreadsheet, "Sheet1")

# Each rule includes the operator and length values
if length_rules != [] do
  rule = hd(length_rules)
  IO.puts("Text length must be #{rule.operator} #{rule.length1}")
end
```

#### Custom Formula Validations

```elixir
# Get all custom formula validations
{:ok, formula_rules} = DataValidation.get_custom_validations(spreadsheet, "Sheet1")

# Each rule includes the custom formula
if formula_rules != [] do
  formula = hd(formula_rules).formula
  IO.puts("Custom validation formula: #{formula}")
end
```

### Filtering Validation Rules by Range

Each getter function accepts an optional cell range parameter that filters the results:

```elixir
# Get only validation rules that apply to range A1:B10
{:ok, rules} = DataValidation.get_data_validations(spreadsheet, "Sheet1", "A1:B10")
```

### Utility Functions

#### Checking If Validations Exist

```elixir
# Check if any validation rules exist
{:ok, has_rules} = DataValidation.has_data_validations(spreadsheet, "Sheet1")
if has_rules do
  IO.puts("Sheet has validation rules")
end

# Check if a specific range has validation rules
{:ok, has_rules_in_range} = DataValidation.has_data_validations(spreadsheet, "Sheet1", "A1:A10")
```

#### Counting Validation Rules

```elixir
# Count all validation rules in a sheet
{:ok, count} = DataValidation.count_data_validations(spreadsheet, "Sheet1")
IO.puts("Sheet has #{count} validation rules")

# Count rules for a specific range
{:ok, range_count} = DataValidation.count_data_validations(spreadsheet, "Sheet1", "A1:A10")
```

## Removing Validation

You can remove all validation rules from a range:

```elixir
DataValidation.remove_data_validation(
  spreadsheet,
  "Sheet1",
  "A1:Z100"
)
```

## Common Parameters

All data validation functions accept these common parameters:

- `allow_blank`: Whether to allow blank/empty values (default: true)
- `error_title`: Title for error message when validation fails (optional)
- `error_message`: Description for error message when validation fails (optional)
- `prompt_title`: Title for input prompt (optional)
- `prompt_message`: Description for input prompt (optional)

All validation rules include these common properties when retrieved:

- `range`: The cell range the validation applies to
- `rule_type`: The type of validation
- `allow_blank`: Whether blank values are allowed
- `show_error_message`: Whether error messages are shown
- `error_title`: The title of the error message (if any)
- `error_message`: The error message text (if any)
- `show_input_message`: Whether input prompts are shown
- `prompt_title`: The title of the input prompt (if any)
- `prompt_message`: The input prompt text (if any)

## Available Operators

The following operators are available for number, date, and text length validation:

- `"between"` - Value must be between two values (inclusive)
- `"notBetween"` - Value must not be between two values
- `"equal"` - Value must equal the specified value
- `"notEqual"` - Value must not equal the specified value
- `"greaterThan"` - Value must be greater than the specified value
- `"lessThan"` - Value must be less than the specified value
- `"greaterThanOrEqual"` - Value must be greater than or equal to the specified value
- `"lessThanOrEqual"` - Value must be less than or equal to the specified value

## Complete Example

Here's a complete example showing various validation rules:

```elixir
alias UmyaSpreadsheet.{DataValidation}

# Create a spreadsheet with validation
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Categories dropdown in column A
DataValidation.add_list_validation(
  spreadsheet,
  "Sheet1",
  "A2:A100",
  ["Product", "Service", "Subscription"],
  true,
  "Invalid Category",
  "Please select a valid category",
  "Category",
  "Select a category for this item"
)

# Price in column B (must be positive)
DataValidation.add_number_validation(
  spreadsheet,
  "Sheet1",
  "B2:B100",
  "greaterThan",
  0.0,
  nil,
  true,
  "Invalid Price",
  "Price must be greater than 0",
  "Price",
  "Enter the item price"
)

# Release date in column C (must be in the past)
DataValidation.add_date_validation(
  spreadsheet,
  "Sheet1",
  "C2:C100",
  "lessThanOrEqual",
  Date.utc_today(),  # Date struct works directly
  nil,
  true,
  "Invalid Date",
  "Release date must be today or in the past",
  "Release Date",
  "Enter when the item was released"
)

# Description in column D (between 10 and 200 chars)
DataValidation.add_text_length_validation(
  spreadsheet,
  "Sheet1",
  "D2:D100",
  "between",
  10,
  200,
  true,
  "Invalid Description",
  "Description must be between 10 and 200 characters",
  "Description",
  "Enter a description (10-200 characters)"
)

# SKU in column E (must match pattern using custom formula)
DataValidation.add_custom_validation(
  spreadsheet,
  "Sheet1",
  "E2:E100",
  "AND(LEN(E2)=8,LEFT(E2,3)=\"PRD\")",
  true,
  "Invalid SKU",
  "SKU must be 8 characters and start with PRD",
  "SKU",
  "Enter a SKU (format: PRDxxxxx)"
)

# Add headers
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Category")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Price")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C1", "Release Date")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D1", "Description")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "E1", "SKU")

# Make headers bold
UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1:E1", true)

# Save the file
UmyaSpreadsheet.write_file(spreadsheet, "product_catalog.xlsx")
```

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

## Tips for Data Validation

1. **Error Messages**: Provide clear error messages to guide users when they enter invalid data.
2. **Input Prompts**: Use prompt titles and messages to provide input instructions.
3. **Allow Blank**: Set to `false` to require input in cells.
4. **Custom Formulas**: For complex validation, use custom formulas. Note that formulas should not include the `=` sign at the beginning.
5. **Testing**: Always test your validation rules by opening the Excel file manually to ensure they work as expected.
6. **Use Date Structs**: For date validation, you can use either ISO date strings or Elixir Date structs.
7. **Check Validation First**: Use the utility functions to check if validation exists before attempting to add or modify rules.
8. **Pattern Matching**: Use pattern matching to efficiently check if validation rules exist (e.g., `{:ok, [_ | _]} = DataValidation.get_data_validations(...)`)
