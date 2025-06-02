# Data Validation Guide

This guide explains how to use data validation in UmyaSpreadsheet to restrict the type of data or values that users can enter into a cell.

## Introduction

Data validation is a powerful feature in Excel that helps ensure data integrity by controlling what values can be entered into cells. UmyaSpreadsheet provides comprehensive support for creating and managing various types of data validation rules.

## Types of Data Validation

UmyaSpreadsheet supports all standard Excel validation types:

- **List Validations**: Restrict input to a list of predefined options (dropdown)
- **Number Validations**: Restrict input to numbers within specified constraints
- **Date Validations**: Restrict input to dates within specified constraints
- **Text Length Validations**: Restrict input based on text length
- **Custom Formula Validations**: Restrict input using custom Excel formulas

## Adding Data Validation Rules

### List Validation (Dropdown)

```elixir
# Create a dropdown list with predefined options
DataValidation.add_list_validation(
  spreadsheet,
  "Sheet1",
  "A1:A10",
  ["Red", "Green", "Blue"],
  true,  # Allow blank values
  "Invalid Color",  # Error title
  "Please select a color from the list",  # Error message
  "Color Selection",  # Input prompt title
  "Select a color from the list"  # Input prompt message
)
```

### Number Validation

```elixir
# Restrict input to numbers between 1 and 10
DataValidation.add_number_validation(
  spreadsheet,
  "Sheet1",
  "B1:B10",
  "between",  # Operator
  1.0,  # Minimum value
  10.0,  # Maximum value
  true,  # Allow blank values
  "Invalid Number",  # Error title
  "Please enter a number between 1 and 10",  # Error message
  "Number Input",  # Input prompt title
  "Enter a number between 1 and 10"  # Input prompt message
)
```

### Date Validation

Date validation can be set up using either ISO date strings or Elixir Date structs:

```elixir
# Using ISO date strings
DataValidation.add_date_validation(
  spreadsheet,
  "Sheet1",
  "C1:C10",
  "between",  # Operator
  "2023-01-01",  # Start date
  "2023-12-31",  # End date
  true,  # Allow blank values
  "Invalid Date",  # Error title
  "Please enter a date in 2023",  # Error message
  "Date Input",  # Input prompt title
  "Enter a date in 2023"  # Input prompt message
)

# Using Elixir Date structs
DataValidation.add_date_validation(
  spreadsheet,
  "Sheet1",
  "D1:D10",
  "greaterThanOrEqual",  # Operator
  Date.utc_today(),  # Today's date
  nil,  # No second date needed
  true,  # Allow blank values
  "Invalid Date",  # Error title
  "Please enter today's date or later"  # Error message
)
```

### Text Length Validation

```elixir
# Restrict text length to maximum 50 characters
DataValidation.add_text_length_validation(
  spreadsheet,
  "Sheet1",
  "E1:E10",
  "lessThanOrEqual",  # Operator
  50,  # Maximum length
  nil,  # No second length needed
  true,  # Allow blank values
  "Text Too Long",  # Error title
  "Please enter text with 50 characters or less"  # Error message
)
```

### Custom Formula Validation

```elixir
# Custom validation to ensure value is divisible by 3
DataValidation.add_custom_validation(
  spreadsheet,
  "Sheet1",
  "F1:F10",
  "MOD(F1,3)=0",  # Formula without = sign
  true,  # Allow blank values
  "Invalid Value",  # Error title
  "Please enter a value divisible by 3"  # Error message
)
```

## Removing Data Validation

To remove validation from a range of cells:

```elixir
DataValidation.remove_data_validation(
  spreadsheet,
  "Sheet1",
  "A1:A10"
)
```

## Retrieving Validation Rules

UmyaSpreadsheet provides functions to retrieve and inspect existing data validation rules. See the [Data Validation Getters Guide](./data_validation_getters.md) for detailed information.

## Common Parameters

All data validation functions accept these common parameters:

- `allow_blank`: Whether to allow blank/empty values (default: true)
- `error_title`: Title for error message when validation fails (optional)
- `error_message`: Description for error message when validation fails (optional)
- `prompt_title`: Title for input prompt (optional)
- `prompt_message`: Description for input prompt (optional)

## Tips for Effective Data Validation

1. **Use clear error messages**: Provide helpful guidance when users enter invalid data
2. **Provide input prompts**: Help users understand what input is expected
3. **Use the right validation type**: Choose the most appropriate validation type for your data
4. **Allow blanks when appropriate**: Consider whether blank values should be permitted

This guide covers the basics of applying data validation to your Excel spreadsheets using UmyaSpreadsheet. For examples of retrieving validation rules, see the [Data Validation Getters Guide](./data_validation_getters.md).
