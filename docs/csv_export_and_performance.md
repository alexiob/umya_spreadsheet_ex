# CSV Export and Performance Options

This document covers the CSV export functionality and the performance-optimized writer options in UmyaSpreadsheet.

## CSV Export

The CSV export feature allows you to export individual sheets from an Excel file to CSV format for easier data interchange.

### Usage

```elixir
# Read an existing spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.read("sales_data.xlsx")

# Export Sheet1 to a CSV file
:ok = UmyaSpreadsheet.write_csv(spreadsheet, "Sheet1", "sales_data.csv")
```

### Options and Limitations

- Each export creates a single CSV file from a single sheet
- The export uses default CSV settings (comma-separated, quoted strings when needed)
- All cell values will be represented as strings in the CSV output
- Formatting and styling is not preserved in the CSV export

## Light Writer Functions

These functions offer memory-efficient alternatives to the standard `write` functions, which is especially useful for large spreadsheets.

### Available Light Writer Functions

1. `write_light/2` - Writes a spreadsheet to a file using less memory
2. `write_with_password_light/3` - Writes a password-protected spreadsheet using less memory

### Usage

```elixir
# Create a large spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add lots of data...
# ...

# Use light writer to save with less memory usage
:ok = UmyaSpreadsheet.write_light(spreadsheet, "large_spreadsheet.xlsx")

# Or with password protection
:ok = UmyaSpreadsheet.write_with_password_light(spreadsheet, "protected_large.xlsx", "secret123")
```

### When to Use Light Writers

- When working with very large spreadsheets (thousands of rows/columns)
- When running in memory-constrained environments
- When maximum performance is needed and some advanced features aren't required

### Limitations of Light Writers

The light writer functions have some trade-offs to achieve their memory efficiency:

- May not preserve all complex cell styles
- Some advanced chart features might be simplified
- Drawing and image handling might be slightly different

## Creating Empty Spreadsheets

The `new_empty/0` function allows you to create a spreadsheet without any default worksheets, which can be useful when you want to entirely customize your spreadsheet structure.

```elixir
# Create an empty spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new_empty()

# Add your own sheet
:ok = UmyaSpreadsheet.add_sheet(spreadsheet, "CustomSheet")

# Now work with your custom sheet
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "CustomSheet", "A1", "Custom data")
```

## Performance Recommendations

For the best performance when working with large Excel files:

1. Use `lazy_read/1` instead of `read/1` when you only need to access certain sheets
2. Use `write_light/2` when saving large files
3. Use `new_empty/0` when you need to create highly customized spreadsheets
4. Remove unnecessary sheets before saving large files
5. Consider exporting individual sheets to CSV if the recipient only needs the data, not the formatting
