# Auto Filters

This guide covers auto filter functionality in UmyaSpreadsheet, allowing you to add Excel-style filtering to your spreadsheet data.

## What Are Auto Filters?

Auto filters add dropdown menus to column headers in your spreadsheet, allowing users to filter and sort data directly in Excel. When a user opens a spreadsheet with auto filters:

1. Filter buttons appear in the header row
2. Clicking a filter button shows options like:
   - Sort A to Z / Sort Z to A
   - Filter by specific values
   - Filter by conditions (greater than, less than, etc.)
   - Search for specific text

Auto filters make large datasets more manageable by allowing users to focus on specific subsets of data.

## Adding Auto Filters

To add an auto filter to a range of data:

```elixir
alias UmyaSpreadsheet

# Create a new spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add some header data
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Name")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Department")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C1", "Salary")

# Add some rows of data
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "John")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", "IT")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C2", "75000")

UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "Sarah")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B3", "HR")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C3", "65000")

# Add an auto filter to the data range
UmyaSpreadsheet.set_auto_filter(spreadsheet, "Sheet1", "A1:C3")

# Save the spreadsheet
UmyaSpreadsheet.write(spreadsheet, "employee_data.xlsx")
```

When this spreadsheet is opened in Excel, filter buttons will appear in cells A1, B1, and C1, allowing users to filter and sort the data.

## Checking If Auto Filters Exist

You can check if a worksheet already has auto filters:

```elixir
# Check if the worksheet has an auto filter
case UmyaSpreadsheet.has_auto_filter(spreadsheet, "Sheet1") do
  true -> IO.puts("Sheet1 has an auto filter")
  false -> IO.puts("Sheet1 does not have an auto filter")
  {:error, reason} -> IO.puts("Error checking auto filter: #{reason}")
end
```

## Getting the Auto Filter Range

To retrieve the range of cells that have auto filters applied:

```elixir
# Get the range of the auto filter
case UmyaSpreadsheet.get_auto_filter_range(spreadsheet, "Sheet1") do
  nil -> IO.puts("No auto filter is set on Sheet1")
  range when is_binary(range) -> IO.puts("Auto filter range: #{range}")
  {:error, reason} -> IO.puts("Error getting auto filter range: #{reason}")
end
```

## Removing Auto Filters

To remove auto filters from a worksheet:

```elixir
# Remove the auto filter from Sheet1
UmyaSpreadsheet.remove_auto_filter(spreadsheet, "Sheet1")
```

## Best Practices

1. **Include Headers**: Always include clear column headers in the first row of your auto filter range
2. **Consistent Data Types**: For best filtering results, keep data types consistent within columns
3. **Appropriate Range**: Ensure your filter range includes all data that should be filterable, but not empty rows
4. **Format Headers**: Consider formatting header cells (bold, background color) to make them stand out
5. **Test Filters**: Open your spreadsheet in Excel to verify filters work as expected

## Example: Financial Data with Auto Filters

Here's a more complex example showing how to create a financial report with auto filters:

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Set up headers with formatting
headers = ["Date", "Transaction", "Category", "Amount", "Balance"]
Enum.with_index(headers, fn header, idx ->
  col = Enum.at(~w(A B C D E), idx)
  UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "#{col}1", header)
  UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "#{col}1", true)
  UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "#{col}1", "DDEBF7")
end)

# Add some transaction data
transactions = [
  ["2025-05-01", "Salary", "Income", 5000, 5000],
  ["2025-05-02", "Rent", "Housing", -1200, 3800],
  ["2025-05-03", "Groceries", "Food", -150, 3650],
  ["2025-05-04", "Restaurant", "Food", -75, 3575],
  ["2025-05-05", "Gas", "Transportation", -40, 3535],
  ["2025-05-06", "Internet", "Utilities", -60, 3475],
  ["2025-05-07", "Movie Tickets", "Entertainment", -30, 3445]
]

Enum.with_index(transactions, fn transaction, idx ->
  row = idx + 2 # Start from row 2
  Enum.with_index(transaction, fn value, col_idx ->
    col = Enum.at(~w(A B C D E), col_idx)
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "#{col}#{row}", value)

    # Format amounts and balance as currency
    if col_idx >= 3 do
      UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "#{col}#{row}", "$#,##0.00")
    end
  end)
end)

# Apply auto filter to the entire data set
UmyaSpreadsheet.set_auto_filter(spreadsheet, "Sheet1", "A1:E8")

# Save the spreadsheet
UmyaSpreadsheet.write(spreadsheet, "financial_report.xlsx")
```

With this setup, users can:

- Filter transactions by category
- Sort by amount (largest to smallest)
- Filter to see only expenses (negative amounts)
- Filter by date ranges

## Limitations

- The library adds the auto filter capability, but the actual filtering happens in Excel or other spreadsheet applications
- Auto filters are applied to entire columns within the specified range, not partial columns
- Complex filter conditions cannot be pre-set; they must be applied by the user in Excel
