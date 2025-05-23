# Pivot Tables

This guide explains how to use UmyaSpreadsheet's pivot table functionality to summarize and analyze data in your spreadsheets.

## Introduction to Pivot Tables

Pivot tables are powerful tools for data analysis that allow you to:

- Summarize large datasets into a more manageable format
- Rearrange data to show different perspectives
- Create custom calculations and aggregations
- Generate insights through cross-tabulations

UmyaSpreadsheet allows you to create pivot tables programmatically, giving you the ability to build analysis tools and reports without manual Excel operations.

## Creating a Basic Pivot Table

To create a pivot table, you need source data in a tabular format with headers in the first row. Then you can use the `add_pivot_table/9` function:

```elixir
UmyaSpreadsheet.add_pivot_table(
  spreadsheet,
  "PivotSheet",                # Sheet where the pivot table will be placed
  "Sales Analysis",            # Name of the pivot table
  "Data",                      # Sheet containing the source data
  "A1:D100",                   # Range of source data (including headers)
  "A3",                        # Top-left cell for pivot table placement
  [0],                         # Row fields (0-based indices of source columns)
  [1],                         # Column fields (0-based indices of source columns)
  [{2, "sum", "Total Sales"}]  # Data fields with aggregation functions
)
```

This will create a pivot table that:
- Uses the first column (index 0) as row labels
- Uses the second column (index 1) as column labels
- Sums the values from the third column (index 2) at the intersections

## Field Types

A pivot table consists of three main field types:

1. **Row Fields**: These fields appear as row labels (vertically) in the pivot table. Each unique value becomes a row in the pivot table.

2. **Column Fields**: These fields appear as column labels (horizontally) in the pivot table. Each unique value becomes a column.

3. **Data Fields**: These fields contain the values that are calculated at the intersection of row and column fields. They require:
   - A field index (which column to use)
   - An aggregation function (how to aggregate the values)
   - A custom name (what to label the data field)

## Available Aggregation Functions

When setting up data fields, you can specify one of the following aggregation functions:

- `"sum"` - Sum of all values
- `"count"` - Count of all values
- `"average"` - Average (mean) of values
- `"max"` - Maximum value
- `"min"` - Minimum value
- `"product"` - Product of all values (multiplication)
- `"count_nums"` - Count of numeric values only
- `"stddev"` - Standard deviation of a sample
- `"stddevp"` - Standard deviation of a population
- `"var"` - Variance of a sample
- `"varp"` - Variance of a population

For example, to create a pivot table with average values:

```elixir
UmyaSpreadsheet.add_pivot_table(
  spreadsheet,
  "PivotSheet",
  "Average Analysis",
  "Data",
  "A1:D100",
  "A3",
  [0],
  [1],
  [{2, "average", "Average Sales"}]
)
```

## Multiple Data Fields

You can include multiple data fields in a single pivot table:

```elixir
UmyaSpreadsheet.add_pivot_table(
  spreadsheet,
  "PivotSheet",
  "Complete Analysis",
  "Data",
  "A1:D100",
  "A3",
  [0],
  [1],
  [
    {2, "sum", "Total Sales"},
    {2, "average", "Average Sales"},
    {2, "max", "Max Sale"}
  ]
)
```

This creates a pivot table with three data fields, all using the same source column but with different aggregation functions.

## Managing Pivot Tables

UmyaSpreadsheet provides several functions to manage pivot tables:

### Checking for Pivot Tables

```elixir
# Check if a sheet has any pivot tables
if UmyaSpreadsheet.has_pivot_tables?(spreadsheet, "PivotSheet") do
  # Do something with sheets that have pivot tables
end

# Count pivot tables in a sheet
count = UmyaSpreadsheet.count_pivot_tables(spreadsheet, "PivotSheet")
```

### Refreshing Pivot Tables

When the source data changes, you need to refresh pivot tables to update them:

```elixir
# Refresh all pivot tables in the spreadsheet
UmyaSpreadsheet.refresh_all_pivot_tables(spreadsheet)
```

### Removing Pivot Tables

You can remove a specific pivot table by its name:

```elixir
# Remove a pivot table
UmyaSpreadsheet.remove_pivot_table(spreadsheet, "PivotSheet", "Sales Analysis")
```

## Complete Example

Here's a complete example showing how to create a spreadsheet with both source data and a pivot table:

```elixir
# Create a new spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Create headers in Sheet1
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Region")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Product")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C1", "Sales")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D1", "Date")

# Add some data
regions = ["North", "South", "East", "West"]
products = ["Apples", "Oranges", "Bananas", "Grapes"]

# Generate random sales data
for row <- 2..50 do
  region = Enum.random(regions)
  product = Enum.random(products)
  sales = :rand.uniform(10000) + 5000
  date = Date.add(Date.utc_today(), -:rand.uniform(365))

  UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A#{row}", region)
  UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B#{row}", product)
  UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C#{row}", "#{sales}")
  UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D#{row}", Date.to_iso8601(date))
end

# Create a sheet for the pivot table
UmyaSpreadsheet.add_sheet(spreadsheet, "Sales Analysis")

# Create region-by-product pivot table
UmyaSpreadsheet.add_pivot_table(
  spreadsheet,
  "Sales Analysis",
  "Region by Product",
  "Sheet1",
  "A1:D50",
  "A3",
  [0], # Region as rows
  [1], # Product as columns
  [{2, "sum", "Total Sales"}] # Sum of sales
)

# Create a second pivot table showing averages
UmyaSpreadsheet.add_pivot_table(
  spreadsheet,
  "Sales Analysis",
  "Product Analysis",
  "Sheet1",
  "A1:D50",
  "A20",
  [1], # Product as rows
  [],  # No column fields
  [
    {2, "sum", "Total Sales"},
    {2, "average", "Average Sale"},
    {2, "count", "Number of Sales"}
  ]
)

# Save the spreadsheet
UmyaSpreadsheet.write(spreadsheet, "sales_analysis.xlsx")
```

## Limitations

- Formatting options for pivot tables (like styles, number formats, etc.) are currently limited
- The pivot table cache is created when the file is saved, but isn't automatically updated when source data changes
- The `refresh_all_pivot_tables/1` function currently has limited functionality in the backend implementation and may not fully refresh the pivot table data in all cases
- For best results, re-create pivot tables after significant source data changes, or ensure you call `refresh_all_pivot_tables/1` before saving the spreadsheet
- The pivot tables will be properly displayed and functional when the file is opened in Excel or other compatible applications

## Excel Compatibility

Pivot tables created with UmyaSpreadsheet are compatible with Microsoft Excel and other spreadsheet applications that support the XLSX format. When opening the file in Excel, you'll be able to interact with the pivot tables normally, including:

- Refreshing the data
- Modifying the layout
- Changing summary functions
- Adding filters
- Creating pivot charts
