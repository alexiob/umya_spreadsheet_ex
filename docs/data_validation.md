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
- **Date Validation**: Restricts input to dates within specific ranges
- **Text Length Validation**: Restricts input to text with specific character count requirements
- **Custom Formula Validation**: Allows advanced validation using Excel formulas

## List Validation (Dropdown Lists)

Create dropdown lists to restrict input to a predefined set of options:

```elixir
# Create a dropdown with simple options
UmyaSpreadsheet.add_list_validation(
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

## Number Validation

Restrict input to numbers that meet specific criteria:

```elixir
# Allow only numbers between 1 and 100
UmyaSpreadsheet.add_number_validation(
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
UmyaSpreadsheet.add_number_validation(
  spreadsheet,
  "Sheet1",
  "C1:C10",
  "greaterThan",
  0.0
)
```

## Date Validation

Restrict input to dates that meet specific criteria:

```elixir
# Allow only dates in the current year
UmyaSpreadsheet.add_date_validation(
  spreadsheet,
  "Sheet1",
  "D1:D10",                # Cell range
  "between",               # Operator
  "#{Date.utc_today().year}-01-01", # Start date
  "#{Date.utc_today().year}-12-31", # End date
  true,                    # Allow blank (optional)
  "Invalid Date",          # Error title (optional)
  "Please enter a date in #{Date.utc_today().year}", # Error message (optional)
  "Date Input",            # Prompt title (optional)
  "Enter a date in #{Date.utc_today().year}"         # Prompt message (optional)
)

# Allow only future dates
UmyaSpreadsheet.add_date_validation(
  spreadsheet,
  "Sheet1",
  "E1:E10",
  "greaterThanOrEqual",
  Date.utc_today() |> Date.to_iso8601()
)
```

## Text Length Validation

Restrict input to text with specific length requirements:

```elixir
# Limit input to at most 280 characters (like a tweet)
UmyaSpreadsheet.add_text_length_validation(
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
UmyaSpreadsheet.add_text_length_validation(
  spreadsheet,
  "Sheet1",
  "G1:G10",
  "greaterThanOrEqual",
  8
)
```

## Custom Formula Validation

Create advanced validation rules using Excel formulas:

```elixir
# Only allow values divisible by 5
UmyaSpreadsheet.add_custom_validation(
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

## Removing Validation

You can remove all validation rules from a range:

```elixir
UmyaSpreadsheet.remove_data_validation(
  spreadsheet,
  "Sheet1",
  "A1:Z100"
)
```

## Complete Example

Here's a complete example showing various validation rules:

```elixir
# Create a spreadsheet with validation
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Categories dropdown in column A
UmyaSpreadsheet.add_list_validation(
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
UmyaSpreadsheet.add_number_validation(
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
UmyaSpreadsheet.add_date_validation(
  spreadsheet,
  "Sheet1",
  "C2:C100",
  "lessThanOrEqual",
  Date.utc_today() |> Date.to_iso8601(),
  nil,
  true,
  "Invalid Date",
  "Release date must be today or in the past",
  "Release Date",
  "Enter when the item was released"
)

# Description in column D (between 10 and 200 chars)
UmyaSpreadsheet.add_text_length_validation(
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
UmyaSpreadsheet.add_custom_validation(
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

## Tips for Data Validation

1. **Error Messages**: Provide clear error messages to guide users when they enter invalid data.
2. **Input Prompts**: Use prompt titles and messages to provide input instructions.
3. **Allow Blank**: Set to `false` to require input in cells.
4. **Custom Formulas**: For complex validation, use custom formulas. Note that formulas should not include the `=` sign at the beginning.
5. **Testing**: Always test your validation rules by opening the Excel file manually to ensure they work as expected.
