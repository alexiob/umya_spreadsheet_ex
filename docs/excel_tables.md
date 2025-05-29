# Excel Tables Guide

Excel tables provide a powerful way to organize, format, and analyze data in structured formats. UmyaSpreadsheet provides comprehensive support for creating, styling, and managing Excel tables with all their advanced features.

## Table of Contents

- [Overview](#overview)
- [Creating Tables](#creating-tables)
- [Table Management](#table-management)
- [Table Styling](#table-styling)
- [Column Management](#column-management)
- [Totals Row](#totals-row)
- [Complete Examples](#complete-examples)
- [Best Practices](#best-practices)
- [Error Handling](#error-handling)

## Overview

Excel tables offer several advantages over regular cell ranges:

- **Structured References**: Use column names in formulas instead of cell references
- **Automatic Formatting**: Consistent styling with built-in table themes
- **Dynamic Ranges**: Tables automatically expand when you add data
- **Built-in Filtering**: Column headers automatically include filter dropdowns
- **Totals Row**: Easy calculations with common functions (SUM, AVERAGE, COUNT, etc.)
- **Professional Appearance**: Clean, organized presentation of data

All table functions in UmyaSpreadsheet return tuples for consistent error handling:

- `{:ok, result}` for successful operations
- `{:error, reason}` for failures

## Creating Tables

### Basic Table Creation

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add some sample data first
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Name")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Department")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C1", "Salary")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D1", "Start Date")

UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "John Doe")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", "Engineering")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C2", 75000)
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D2", "2020-01-15")

UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "Jane Smith")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B3", "Marketing")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C3", 65000)
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D3", "2019-03-20")

# Create a table
{:ok, :ok} = UmyaSpreadsheet.add_table(
  spreadsheet,
  "Sheet1",              # Worksheet name
  "EmployeeTable",       # Table name (must be unique)
  "Employee Data",       # Display name
  "A1",                  # Top-left corner
  "D3",                  # Bottom-right corner
  ["Name", "Department", "Salary", "Start Date"], # Column headers
  false                  # Show totals row (true/false)
)
```

### Table with Totals Row

```elixir
# Create a table with a totals row enabled
{:ok, :ok} = UmyaSpreadsheet.add_table(
  spreadsheet,
  "Sheet1",
  "SalesTable",
  "Sales Data",
  "A1",
  "E10",
  ["Region", "Product", "Sales", "Quantity", "Revenue"],
  true  # Enable totals row
)
```

## Table Management

### Checking for Tables

```elixir
# Check if a worksheet has any tables
{:ok, has_tables} = UmyaSpreadsheet.has_tables?(spreadsheet, "Sheet1")
# => {:ok, true} or {:ok, false}

# Count tables on a worksheet
{:ok, count} = UmyaSpreadsheet.count_tables(spreadsheet, "Sheet1")
# => {:ok, 2}
```

### Retrieving Table Information

```elixir
# Get all tables from a worksheet
{:ok, tables} = UmyaSpreadsheet.get_tables(spreadsheet, "Sheet1")

# Each table is a map with detailed information
[table | _] = tables
table["name"]         # => "EmployeeTable"
table["display_name"] # => "Employee Data"
table["columns"]      # => [list of column definitions]
table["ref"]          # => "A1:D3" (table range)
table["totals_row_shown"] # => true/false
```

### Table Column Information

```elixir
{:ok, [table]} = UmyaSpreadsheet.get_tables(spreadsheet, "Sheet1")
columns = table["columns"]

# Each column has properties:
Enum.each(columns, fn column ->
  IO.puts("Column: #{column["name"]}")
  IO.puts("ID: #{column["id"]}")
  IO.puts("Has Totals: #{column["totals_row_label"]}")
  IO.puts("Totals Function: #{column["totals_row_function"]}")
end)
```

### Inspecting Specific Tables

UmyaSpreadsheet provides dedicated getter functions for inspecting specific table properties:

```elixir
# Get detailed information about a specific table
{:ok, table_info} = UmyaSpreadsheet.get_table(spreadsheet, "Sheet1", "EmployeeTable")
# => {:ok, %{
#      "name" => "EmployeeTable",
#      "display_name" => "Employee Data",
#      "ref" => "A1:D3",
#      "totals_row_shown" => true,
#      "columns" => [...]
#    }}

# Get just the table style information
{:ok, style_info} = UmyaSpreadsheet.get_table_style(spreadsheet, "Sheet1", "EmployeeTable")
# => {:ok, %{
#      "name" => "TableStyleMedium9",
#      "show_first_column" => true,
#      "show_last_column" => false,
#      "show_row_stripes" => true,
#      "show_column_stripes" => false
#    }}

# Get table column definitions
{:ok, columns} = UmyaSpreadsheet.get_table_columns(spreadsheet, "Sheet1", "EmployeeTable")
# => {:ok, [
#      %{"id" => 1, "name" => "Name", "totals_row_function" => "none"},
#      %{"id" => 2, "name" => "Department", "totals_row_function" => "none"},
#      %{"id" => 3, "name" => "Salary", "totals_row_function" => "sum"}
#    ]}

# Check if totals row is visible
{:ok, has_totals} = UmyaSpreadsheet.get_table_totals_row(spreadsheet, "Sheet1", "EmployeeTable")
# => {:ok, true} or {:ok, false}
```

### Practical Table Inspection Example

```elixir
def inspect_table_details(spreadsheet, sheet_name, table_name) do
  with {:ok, table} <- UmyaSpreadsheet.get_table(spreadsheet, sheet_name, table_name),
       {:ok, style} <- UmyaSpreadsheet.get_table_style(spreadsheet, sheet_name, table_name),
       {:ok, columns} <- UmyaSpreadsheet.get_table_columns(spreadsheet, sheet_name, table_name),
       {:ok, has_totals} <- UmyaSpreadsheet.get_table_totals_row(spreadsheet, sheet_name, table_name) do

    IO.puts("=== Table Details ===")
    IO.puts("Name: #{table["name"]}")
    IO.puts("Display Name: #{table["display_name"]}")
    IO.puts("Range: #{table["ref"]}")
    IO.puts("Has Totals Row: #{has_totals}")

    IO.puts("\n=== Style Information ===")
    IO.puts("Style: #{style["name"]}")
    IO.puts("First Column Highlighted: #{style["show_first_column"]}")
    IO.puts("Last Column Highlighted: #{style["show_last_column"]}")
    IO.puts("Row Stripes: #{style["show_row_stripes"]}")
    IO.puts("Column Stripes: #{style["show_column_stripes"]}")

    IO.puts("\n=== Columns ===")
    Enum.each(columns, fn column ->
      IO.puts("- #{column["name"]} (ID: #{column["id"]}, Function: #{column["totals_row_function"]})")
    end)

    {:ok, :inspection_complete}
  else
    {:error, reason} ->
      IO.puts("Error inspecting table: #{reason}")
      {:error, reason}
  end
end

# Usage
inspect_table_details(spreadsheet, "Sheet1", "EmployeeTable")
```

### Removing Tables

```elixir
# Remove a table by name
{:ok, :ok} = UmyaSpreadsheet.remove_table(spreadsheet, "Sheet1", "EmployeeTable")

# Verify removal
{:ok, false} = UmyaSpreadsheet.has_tables?(spreadsheet, "Sheet1")
{:ok, 0} = UmyaSpreadsheet.count_tables(spreadsheet, "Sheet1")
```

## Table Styling

### Available Table Styles

Excel provides many built-in table styles. Common styles include:

- `"TableStyleLight1"` through `"TableStyleLight21"`
- `"TableStyleMedium1"` through `"TableStyleMedium28"`
- `"TableStyleDark1"` through `"TableStyleDark11"`

### Applying Table Styles

```elixir
# Apply a table style with options
{:ok, :ok} = UmyaSpreadsheet.set_table_style(
  spreadsheet,
  "Sheet1",
  "EmployeeTable",
  "TableStyleMedium9",   # Style name
  true,                  # Show first column (bold formatting)
  false,                 # Show last column (bold formatting)
  true,                  # Show banded rows (alternating colors)
  false                  # Show banded columns (alternating colors)
)
```

### Style Options Explained

- **Show First Column**: Applies bold formatting to the first column
- **Show Last Column**: Applies bold formatting to the last column
- **Show Banded Rows**: Alternates row background colors for better readability
- **Show Banded Columns**: Alternates column background colors

### Popular Style Combinations

```elixir
# Professional blue style with banded rows
{:ok, :ok} = UmyaSpreadsheet.set_table_style(
  spreadsheet, "Sheet1", "DataTable", "TableStyleMedium2",
  true, false, true, false
)

# Light style with column banding
{:ok, :ok} = UmyaSpreadsheet.set_table_style(
  spreadsheet, "Sheet1", "DataTable", "TableStyleLight16",
  false, false, false, true
)

# Dark style for emphasis
{:ok, :ok} = UmyaSpreadsheet.set_table_style(
  spreadsheet, "Sheet1", "DataTable", "TableStyleDark3",
  true, true, true, false
)
```

### Removing Table Styles

```elixir
# Remove all styling from a table (keeps table structure)
{:ok, :ok} = UmyaSpreadsheet.remove_table_style(spreadsheet, "Sheet1", "EmployeeTable")
```

## Column Management

### Adding New Columns

```elixir
# Add a new column to an existing table
{:ok, :ok} = UmyaSpreadsheet.add_table_column(
  spreadsheet,
  "Sheet1",
  "EmployeeTable",
  "Email",               # New column name
  "count",               # Totals row function
  "Total Employees"      # Totals row label
)
```

### Available Totals Row Functions

- `"none"` - No calculation
- `"sum"` - Sum of values
- `"average"` - Average of values
- `"count"` - Count of non-empty cells
- `"countNums"` - Count of numeric cells
- `"max"` - Maximum value
- `"min"` - Minimum value
- `"stdDev"` - Standard deviation
- `"var"` - Variance

### Modifying Existing Columns

```elixir
# Modify an existing table column
{:ok, :ok} = UmyaSpreadsheet.modify_table_column(
  spreadsheet,
  "Sheet1",
  "EmployeeTable",
  "Salary",              # Existing column name
  "Annual Salary",       # New column name
  "average",             # New totals function
  "Average Salary"       # New totals label
)
```

## Totals Row

### Enabling/Disabling Totals Row

```elixir
# Enable totals row
{:ok, :ok} = UmyaSpreadsheet.set_table_totals_row(spreadsheet, "Sheet1", "EmployeeTable", true)

# Disable totals row
{:ok, :ok} = UmyaSpreadsheet.set_table_totals_row(spreadsheet, "Sheet1", "EmployeeTable", false)
```

The totals row automatically calculates values based on the functions assigned to each column when you create or modify the table.

## Complete Examples

### Employee Management System

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Setup employee data
employee_data = [
  ["Name", "Department", "Salary", "Start Date", "Status"],
  ["John Doe", "Engineering", 75000, "2020-01-15", "Active"],
  ["Jane Smith", "Marketing", 65000, "2019-03-20", "Active"],
  ["Bob Johnson", "Engineering", 80000, "2021-06-10", "Active"],
  ["Alice Brown", "Sales", 55000, "2022-02-01", "Active"],
  ["Charlie Davis", "Marketing", 62000, "2020-11-15", "Active"]
]

# Add data to worksheet
Enum.with_index(employee_data, 1)
|> Enum.each(fn {row, row_index} ->
  Enum.with_index(row, 1)
  |> Enum.each(fn {value, col_index} ->
    col = <<(?A + col_index - 1)>>
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "#{col}#{row_index}", value)
  end)
end)

# Create employee table
{:ok, :ok} = UmyaSpreadsheet.add_table(
  spreadsheet,
  "Sheet1",
  "EmployeeTable",
  "Employee Directory",
  "A1",
  "E6",
  ["Name", "Department", "Salary", "Start Date", "Status"],
  true
)

# Apply professional styling
{:ok, :ok} = UmyaSpreadsheet.set_table_style(
  spreadsheet,
  "Sheet1",
  "EmployeeTable",
  "TableStyleMedium9",
  true,   # Highlight first column (names)
  false,  # Don't highlight last column
  true,   # Use banded rows for readability
  false   # Don't use banded columns
)

# Add calculated column
{:ok, :ok} = UmyaSpreadsheet.add_table_column(
  spreadsheet,
  "Sheet1",
  "EmployeeTable",
  "Years of Service",
  "average",
  "Avg Years"
)

# Modify salary column to show totals
{:ok, :ok} = UmyaSpreadsheet.modify_table_column(
  spreadsheet,
  "Sheet1",
  "EmployeeTable",
  "Salary",
  "Annual Salary",
  "sum",
  "Total Payroll"
)

# Save the file
UmyaSpreadsheet.write(spreadsheet, "employee_directory.xlsx")
```

### Sales Report with Multiple Tables

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Create quarterly sales data
quarterly_data = [
  ["Quarter", "Region", "Sales", "Target", "Achievement %"],
  ["Q1 2024", "North", 125000, 120000, 104.2],
  ["Q1 2024", "South", 98000, 100000, 98.0],
  ["Q1 2024", "East", 156000, 150000, 104.0],
  ["Q1 2024", "West", 142000, 140000, 101.4]
]

# Add quarterly data
Enum.with_index(quarterly_data, 1)
|> Enum.each(fn {row, row_index} ->
  Enum.with_index(row, 1)
  |> Enum.each(fn {value, col_index} ->
    col = <<(?A + col_index - 1)>>
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "#{col}#{row_index}", value)
  end)
end)

# Create quarterly table
{:ok, :ok} = UmyaSpreadsheet.add_table(
  spreadsheet,
  "Sheet1",
  "QuarterlySales",
  "Q1 2024 Sales Report",
  "A1",
  "E5",
  ["Quarter", "Region", "Sales", "Target", "Achievement %"],
  true
)

# Style quarterly table
{:ok, :ok} = UmyaSpreadsheet.set_table_style(
  spreadsheet,
  "Sheet1",
  "QuarterlySales",
  "TableStyleMedium2",
  false, false, true, false
)

# Create summary table in another area
summary_data = [
  ["Metric", "Value"],
  ["Total Sales", 521000],
  ["Total Target", 510000],
  ["Overall Achievement", 102.2],
  ["Number of Regions", 4]
]

# Add summary data starting from row 8
Enum.with_index(summary_data, 8)
|> Enum.each(fn {row, row_index} ->
  Enum.with_index(row, 1)
  |> Enum.each(fn {value, col_index} ->
    col = <<(?A + col_index - 1)>>
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "#{col}#{row_index}", value)
  end)
end)

# Create summary table
{:ok, :ok} = UmyaSpreadsheet.add_table(
  spreadsheet,
  "Sheet1",
  "SummaryTable",
  "Summary Metrics",
  "A8",
  "B12",
  ["Metric", "Value"],
  false
)

# Style summary table differently
{:ok, :ok} = UmyaSpreadsheet.set_table_style(
  spreadsheet,
  "Sheet1",
  "SummaryTable",
  "TableStyleDark3",
  true, true, false, false
)

# Verify both tables exist
{:ok, 2} = UmyaSpreadsheet.count_tables(spreadsheet, "Sheet1")
{:ok, tables} = UmyaSpreadsheet.get_tables(spreadsheet, "Sheet1")

Enum.each(tables, fn table ->
  IO.puts("Table: #{table["display_name"]} (#{table["name"]})")
  IO.puts("Range: #{table["ref"]}")
  IO.puts("Columns: #{length(table["columns"])}")
end)

UmyaSpreadsheet.write(spreadsheet, "sales_report.xlsx")
```

## Best Practices

### 1. Table Naming

- Use descriptive, unique table names
- Avoid spaces and special characters
- Follow a consistent naming convention (e.g., `EmployeeData`, `SalesQ1`, `ProductCatalog`)

```elixir
# Good naming
{:ok, :ok} = UmyaSpreadsheet.add_table(spreadsheet, "Sheet1", "EmployeeData", "Employee Directory", ...)

# Avoid
{:ok, :ok} = UmyaSpreadsheet.add_table(spreadsheet, "Sheet1", "Table1", "Table", ...)
```

### 2. Data Preparation

- Ensure data is properly formatted before creating tables
- Include appropriate headers
- Use consistent data types in columns

```elixir
# Prepare data first
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Date")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Amount")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "2024-01-15")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", 1500.50)

# Then create table
{:ok, :ok} = UmyaSpreadsheet.add_table(spreadsheet, "Sheet1", "ExpenseTable", ...)
```

### 3. Style Consistency

- Use consistent table styles across your workbook
- Choose styles that match your organization's branding
- Consider readability and printing requirements

```elixir
# Define standard styles for your application
@standard_table_style "TableStyleMedium9"
@summary_table_style "TableStyleDark3"
@data_table_style "TableStyleLight16"
```

### 4. Error Handling

Always handle potential errors when working with tables:

```elixir
case UmyaSpreadsheet.add_table(spreadsheet, "Sheet1", "DataTable", ...) do
  {:ok, :ok} ->
    IO.puts("Table created successfully")

  {:error, reason} ->
    IO.puts("Failed to create table: #{reason}")
    # Handle error appropriately
end
```

### 5. Performance Considerations

- Create tables after adding all data to avoid multiple range adjustments
- Use batch operations when possible
- Consider memory usage for large tables

## Error Handling

### Common Error Scenarios

```elixir
# Sheet not found
{:error, "Sheet not found"} = UmyaSpreadsheet.add_table(spreadsheet, "NonExistent", ...)

# Table already exists
{:error, "Table with this name already exists"} = UmyaSpreadsheet.add_table(spreadsheet, "Sheet1", "ExistingTable", ...)

# Invalid range
{:error, "Invalid cell range"} = UmyaSpreadsheet.add_table(spreadsheet, "Sheet1", "Table", "Display", "Z1", "AA5", ...)

# Table not found
{:error, "Table not found"} = UmyaSpreadsheet.remove_table(spreadsheet, "Sheet1", "NonExistentTable")
```

### Error Recovery Patterns

```elixir
defmodule TableManager do
  def create_table_safely(spreadsheet, sheet, name, display_name, range_start, range_end, columns) do
    case UmyaSpreadsheet.add_table(spreadsheet, sheet, name, display_name, range_start, range_end, columns, false) do
      {:ok, :ok} ->
        {:ok, :table_created}

      {:error, "Table with this name already exists"} ->
        # Try with a modified name
        new_name = "#{name}_#{System.system_time(:second)}"
        UmyaSpreadsheet.add_table(spreadsheet, sheet, new_name, display_name, range_start, range_end, columns, false)

      {:error, reason} ->
        {:error, reason}
    end
  end

  def ensure_table_exists(spreadsheet, sheet, table_name) do
    case UmyaSpreadsheet.has_tables?(spreadsheet, sheet) do
      {:ok, true} ->
        case UmyaSpreadsheet.get_tables(spreadsheet, sheet) do
          {:ok, tables} ->
            case Enum.find(tables, fn table -> table["name"] == table_name end) do
              nil -> {:error, :table_not_found}
              table -> {:ok, table}
            end

          error -> error
        end

      {:ok, false} ->
        {:error, :no_tables}

      error ->
        error
    end
  end
end
```

This comprehensive guide covers all aspects of working with Excel tables in UmyaSpreadsheet. Tables provide a powerful way to organize and present data professionally, and with these tools, you can create sophisticated data presentations programmatically.
