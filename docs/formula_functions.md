# Formula Functions

This guide covers the advanced formula functionality provided by UmyaSpreadsheet, allowing you to work with cell formulas, array formulas, named ranges, and defined names.

## Regular Cell Formulas

Regular cell formulas are the most common type of formulas in Excel. They are executed in a single cell and return a single value.

```elixir
alias UmyaSpreadsheet

# Create a new spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Set some values to use in our formula
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", 10)
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", 20)
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", 30)

# Set a formula in cell B1 that sums the values in A1:A3
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "B1", "SUM(A1:A3)")

# Save the spreadsheet
UmyaSpreadsheet.write(spreadsheet, "formulas.xlsx")
```

Note: When setting formulas, do not include the leading equals sign (`=`). The library adds it automatically.

### Common Formula Examples

Here are examples of various formulas you can use:

```elixir
# Math operations
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "C1", "A1+A2")
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "C2", "A1*A2")
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "C3", "(A1+A2)/2")  # Average

# Statistical functions
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "D1", "AVERAGE(A1:A10)")
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "D2", "MIN(A1:A10)")
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "D3", "MAX(A1:A10)")
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "D4", "COUNT(A1:A10)")
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "D5", "STDEV(A1:A10)")

# Text manipulation
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "E1", "CONCATENATE(F1, \" \", G1)")
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "E2", "LEFT(F1, 3)")
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "E3", "RIGHT(F1, 4)")
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "E4", "LEN(F1)")

# Logical formulas
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "F1", "IF(A1>20, \"High\", \"Low\")")
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "F2", "AND(A1>5, A1<25)")
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "F3", "OR(A1<5, A1>25)")

# Date and time
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "G1", "TODAY()")
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "G2", "NOW()")
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "G3", "YEAR(TODAY())")
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "G4", "MONTH(TODAY())")
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "G5", "DAY(TODAY())")

# Lookup and reference
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "H1", "VLOOKUP(I1, A1:B10, 2, FALSE)")
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "H2", "HLOOKUP(I1, A1:J2, 2, FALSE)")
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "H3", "INDEX(A1:B10, MATCH(I1, A1:A10, 0), 2)")

# Financial functions
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "I1", "PMT(0.05/12, 360, 200000)")  # Monthly payment for a mortgage
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "I2", "FV(0.05/12, 120, -100)")  # Future value of monthly investment
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "I3", "NPV(0.1, B1:B5)")  # Net Present Value
```

## Array Formulas

Array formulas can return multiple values across a range of cells. They are especially useful for performing calculations on arrays of values.

```elixir
# Set an array formula that fills cells C1:C3 with the values from A1:A3
UmyaSpreadsheet.set_array_formula(spreadsheet, "Sheet1", "C1:C3", "A1:A3")

# Set an array formula that calculates the row numbers
UmyaSpreadsheet.set_array_formula(spreadsheet, "Sheet1", "D1:D5", "ROW(1:5)")
```

### Practical Array Formula Examples

Array formulas are powerful for performing calculations across multiple cells:

```elixir
# Create a sample spreadsheet with data
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add a table of sales data
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Product")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Region")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C1", "Sales")

UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "Apples")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", "North")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C2", 5000)

UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "Apples")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B3", "South")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C3", 4200)

UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A4", "Oranges")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B4", "North")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C4", 3800)

UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A5", "Oranges")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B5", "South")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C5", 4100)

# Example 1: Sum-if array formula to calculate total sales by region
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "E1", "Region")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "F1", "Total Sales")

UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "E2", "North")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "E3", "South")

# Array formula to calculate total sales for North region
UmyaSpreadsheet.set_array_formula(
  spreadsheet,
  "Sheet1",
  "F2",
  "SUM(IF(B2:B5=\"North\",C2:C5,0))"
)

# Array formula to calculate total sales for South region
UmyaSpreadsheet.set_array_formula(
  spreadsheet,
  "Sheet1",
  "F3",
  "SUM(IF(B2:B5=\"South\",C2:C5,0))"
)

# Example 2: Array formula to calculate product-wise sales
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "E5", "Product")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "F5", "Total Sales")

UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "E6", "Apples")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "E7", "Oranges")

# Array formula for Apples total sales
UmyaSpreadsheet.set_array_formula(
  spreadsheet,
  "Sheet1",
  "F6",
  "SUM(IF(A2:A5=\"Apples\",C2:C5,0))"
)

# Array formula for Oranges total sales
UmyaSpreadsheet.set_array_formula(
  spreadsheet,
  "Sheet1",
  "F7",
  "SUM(IF(A2:A5=\"Oranges\",C2:C5,0))"
)

# Example 3: Array formula for calculating multiple statistics at once
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "H1", "Sales Statistics")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "H2", "Average")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "H3", "Min")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "H4", "Max")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "H5", "Count")

# Array formula that populates statistics for Sales data
UmyaSpreadsheet.set_array_formula(
  spreadsheet,
  "Sheet1",
  "I2:I5",
  "{AVERAGE(C2:C5);MIN(C2:C5);MAX(C2:C5);COUNT(C2:C5)}"
)
```

## Named Ranges

Named ranges allow you to refer to a range of cells by a custom name, making formulas more readable and easier to maintain.

```elixir
# Create a named range called "Data" that refers to cells A1:A10
UmyaSpreadsheet.create_named_range(spreadsheet, "Data", "Sheet1", "A1:A10")

# Now you can use this named range in a formula
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "B10", "SUM(Data)")
```

### Named Range Examples for Financial Reports

Here's a complete example of using named ranges for creating a financial report:

```elixir
alias UmyaSpreadsheet

# Create a new spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Set up headers and data for a quarterly financial report
headers = ["Q1", "Q2", "Q3", "Q4", "Total"]
revenue_data = [125000, 142000, 158000, 175000]
expenses_data = [95000, 102000, 110000, 118000]
categories = ["Revenue", "Expenses", "Profit"]

# Add headers
Enum.with_index(headers, fn header, idx ->
  UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B#{idx + 1}", header)
end)

# Add category labels
Enum.with_index(categories, fn category, idx ->
  UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A#{idx + 2}", category)
end)

# Add revenue data
Enum.with_index(revenue_data, fn value, idx ->
  UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B#{2 + 0}", revenue_data)
end)

# Add expense data
Enum.with_index(expenses_data, fn value, idx ->
  UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B#{2 + 1}", expenses_data)
end)

# Create named ranges for better formula readability
UmyaSpreadsheet.create_named_range(spreadsheet, "Revenue", "Sheet1", "B2:E2")
UmyaSpreadsheet.create_named_range(spreadsheet, "Expenses", "Sheet1", "B3:E3")

# Calculate quarterly profits using named ranges
UmyaSpreadsheet.set_array_formula(spreadsheet, "Sheet1", "B4:E4", "Revenue-Expenses")

# Calculate totals using named ranges
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "F2", "SUM(Revenue)")
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "F3", "SUM(Expenses)")
UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "F4", "SUM(B4:E4)")

# Add some interesting metrics at the bottom
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A6", "Profit Margin (%)")
UmyaSpreadsheet.create_named_range(spreadsheet, "TotalRevenue", "Sheet1", "F2")
UmyaSpreadsheet.create_named_range(spreadsheet, "TotalProfit", "Sheet1", "F4")

# Use named ranges in the profit margin calculation
UmyaSpreadsheet.set_formula(
  spreadsheet,
  "Sheet1",
  "B6",
  "ROUND((TotalProfit/TotalRevenue)*100, 1)"
)

# Format cells (optional)
UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1:F1", true)
UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A2:A6", true)
UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "F2:F4", true)
UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "B2:F4", "#,##0")
UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "B6", "0.0%")

# Save the report
UmyaSpreadsheet.write(spreadsheet, "financial_report.xlsx")
```

## Defined Names

Defined names are similar to named ranges but can store formulas or constants instead of just cell references.

```elixir
# Create a defined name for a constant value
UmyaSpreadsheet.create_defined_name(spreadsheet, "TaxRate", "0.15")

# Create a defined name for a formula
UmyaSpreadsheet.create_defined_name(spreadsheet, "Subtotal", "SUM(A1:A10)")

# Create a sheet-scoped defined name
UmyaSpreadsheet.create_defined_name(spreadsheet, "Department", "Sales", "Sheet1")
```

Sheet-scoped defined names are only available within the specified sheet.

### Advanced Defined Names Example: Sales Calculator

Here's a more complex example using defined names to create a sales calculator spreadsheet:

```elixir
alias UmyaSpreadsheet

# Create a new spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Create a sheet for our calculator
UmyaSpreadsheet.add_sheet(spreadsheet, "Sales Calculator")

# Set up headers and formatting
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sales Calculator", "A1", "Sales Calculator")
UmyaSpreadsheet.set_font_bold(spreadsheet, "Sales Calculator", "A1", true)
UmyaSpreadsheet.set_font_size(spreadsheet, "Sales Calculator", "A1", 16)

# Input section
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sales Calculator", "A3", "Input Parameters")
UmyaSpreadsheet.set_font_bold(spreadsheet, "Sales Calculator", "A3", true)

# Create input fields
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sales Calculator", "A4", "Product Price")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sales Calculator", "B4", 100)

UmyaSpreadsheet.set_cell_value(spreadsheet, "Sales Calculator", "A5", "Quantity")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sales Calculator", "B5", 5)

UmyaSpreadsheet.set_cell_value(spreadsheet, "Sales Calculator", "A6", "Tax Rate (%)")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sales Calculator", "B6", 15)

UmyaSpreadsheet.set_cell_value(spreadsheet, "Sales Calculator", "A7", "Discount Rate (%)")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sales Calculator", "B7", 10)

# Create defined names for all the parameters (easier to maintain formulas)
UmyaSpreadsheet.create_defined_name(spreadsheet, "Price", "Sales Calculator!B4")
UmyaSpreadsheet.create_defined_name(spreadsheet, "Quantity", "Sales Calculator!B5")
UmyaSpreadsheet.create_defined_name(spreadsheet, "TaxRatePercent", "Sales Calculator!B6")
UmyaSpreadsheet.create_defined_name(spreadsheet, "DiscountRatePercent", "Sales Calculator!B7")

# Create defined names for calculations (constants)
UmyaSpreadsheet.create_defined_name(spreadsheet, "TaxRate", "TaxRatePercent/100")
UmyaSpreadsheet.create_defined_name(spreadsheet, "DiscountRate", "DiscountRatePercent/100")

# Output section
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sales Calculator", "A9", "Calculation Results")
UmyaSpreadsheet.set_font_bold(spreadsheet, "Sales Calculator", "A9", true)

UmyaSpreadsheet.set_cell_value(spreadsheet, "Sales Calculator", "A10", "Subtotal")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sales Calculator", "A11", "Discount Amount")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sales Calculator", "A12", "Net Amount")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sales Calculator", "A13", "Tax Amount")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sales Calculator", "A14", "Total Due")

# Create formulas using defined names
UmyaSpreadsheet.create_defined_name(spreadsheet, "Subtotal", "Price*Quantity")
UmyaSpreadsheet.set_formula(spreadsheet, "Sales Calculator", "B10", "Subtotal")

UmyaSpreadsheet.create_defined_name(spreadsheet, "DiscountAmount", "Subtotal*DiscountRate")
UmyaSpreadsheet.set_formula(spreadsheet, "Sales Calculator", "B11", "DiscountAmount")

UmyaSpreadsheet.create_defined_name(spreadsheet, "NetAmount", "Subtotal-DiscountAmount")
UmyaSpreadsheet.set_formula(spreadsheet, "Sales Calculator", "B12", "NetAmount")

UmyaSpreadsheet.create_defined_name(spreadsheet, "TaxAmount", "NetAmount*TaxRate")
UmyaSpreadsheet.set_formula(spreadsheet, "Sales Calculator", "B13", "TaxAmount")

UmyaSpreadsheet.create_defined_name(spreadsheet, "TotalDue", "NetAmount+TaxAmount")
UmyaSpreadsheet.set_formula(spreadsheet, "Sales Calculator", "B14", "TotalDue")

# Apply currency formatting to amounts
UmyaSpreadsheet.set_number_format(spreadsheet, "Sales Calculator", "B10:B14", "$#,##0.00")

# Add a second sheet with alternative tax rates for comparison
UmyaSpreadsheet.add_sheet(spreadsheet, "Tax Comparison")

# Create a defined name that's scoped just to this sheet
UmyaSpreadsheet.create_defined_name(
  spreadsheet,
  "AlternateTaxRate",
  "0.08",  # 8% tax rate
  "Tax Comparison"
)

UmyaSpreadsheet.set_cell_value(spreadsheet, "Tax Comparison", "A1", "Sales with Alternative Tax Rate")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Tax Comparison", "A3", "Net Amount")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Tax Comparison", "A4", "Alternative Tax (8%)")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Tax Comparison", "A5", "Total with Alternative Tax")

# Use both global and sheet-specific defined names
UmyaSpreadsheet.set_formula(spreadsheet, "Tax Comparison", "B3", "NetAmount")
UmyaSpreadsheet.set_formula(spreadsheet, "Tax Comparison", "B4", "NetAmount*AlternateTaxRate")
UmyaSpreadsheet.set_formula(spreadsheet, "Tax Comparison", "B5", "B3+B4")

# Apply the same currency formatting
UmyaSpreadsheet.set_number_format(spreadsheet, "Tax Comparison", "B3:B5", "$#,##0.00")

# Save the spreadsheet
UmyaSpreadsheet.write(spreadsheet, "sales_calculator.xlsx")
```

## Listing Defined Names

You can retrieve all defined names in a workbook:

```elixir
defined_names = UmyaSpreadsheet.get_defined_names(spreadsheet)
# Returns: [{"Data", "Sheet1!A1:A10"}, {"TaxRate", "0.15"}, ...]

# Display all defined names
Enum.each(defined_names, fn {name, address} ->
  IO.puts("#{name}: #{address}")
end)
```

### Working with Retrieved Defined Names

Here's an example of how you might use the `get_defined_names` function to generate documentation for a spreadsheet:

```elixir
alias UmyaSpreadsheet

# Read an existing spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.read("financial_model.xlsx")

# Get a list of all defined names
defined_names = UmyaSpreadsheet.get_defined_names(spreadsheet)

# Create a documentation sheet
UmyaSpreadsheet.add_sheet(spreadsheet, "Documentation")

# Add headers
UmyaSpreadsheet.set_cell_value(spreadsheet, "Documentation", "A1", "Name")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Documentation", "B1", "Reference/Formula")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Documentation", "C1", "Description")

# Format headers
UmyaSpreadsheet.set_font_bold(spreadsheet, "Documentation", "A1:C1", true)
UmyaSpreadsheet.set_background_color(spreadsheet, "Documentation", "A1:C1", "DDDDDD")

# Add all defined names to the documentation sheet
Enum.with_index(defined_names, fn {name, address}, index ->
  row = index + 2  # Start from row 2 (after headers)

  UmyaSpreadsheet.set_cell_value(spreadsheet, "Documentation", "A#{row}", name)
  UmyaSpreadsheet.set_cell_value(spreadsheet, "Documentation", "B#{row}", address)
  UmyaSpreadsheet.set_cell_value(spreadsheet, "Documentation", "C#{row}", "")  # Empty description for user to fill
end)

# Auto-adjust column widths for better readability
UmyaSpreadsheet.set_column_auto_width(spreadsheet, "Documentation", "A", true)
UmyaSpreadsheet.set_column_auto_width(spreadsheet, "Documentation", "B", true)
UmyaSpreadsheet.set_column_width(spreadsheet, "Documentation", "C", 50.0)  # Wide column for descriptions

# Save the updated spreadsheet
UmyaSpreadsheet.write(spreadsheet, "financial_model_documented.xlsx")
```

Another practical application is to build a workbook analyzer that reports on formula complexity:

```elixir
alias UmyaSpreadsheet

# Function to analyze a spreadsheet for complex formulas
def analyze_spreadsheet_formulas(path) do
  {:ok, spreadsheet} = UmyaSpreadsheet.read(path)
  defined_names = UmyaSpreadsheet.get_defined_names(spreadsheet)

  # Group defined names by their complexity
  {simple_formulas, complex_formulas} = Enum.split_with(defined_names, fn {_name, formula} ->
    # Simple heuristic: formulas with fewer than 3 operators are "simple"
    operator_count = formula
      |> String.graphemes()
      |> Enum.count(fn char -> char in ["+", "-", "*", "/", "=", ">", "<"] end)

    operator_count < 3
  end)

  # Report on findings
  IO.puts("Spreadsheet Formula Analysis: #{path}")
  IO.puts("Total defined names: #{length(defined_names)}")
  IO.puts("Simple formulas: #{length(simple_formulas)}")
  IO.puts("Complex formulas: #{length(complex_formulas)}")

  IO.puts("\nMost complex formulas:")
  complex_formulas
  |> Enum.sort_by(fn {_name, formula} -> String.length(formula) end, :desc)
  |> Enum.take(5)
  |> Enum.each(fn {name, formula} ->
    IO.puts("#{name}: #{if String.length(formula) > 50, do: String.slice(formula, 0, 47) <> "...", else: formula}")
  end)

  # Return analysis results
  %{
    total: length(defined_names),
    simple: length(simple_formulas),
    complex: length(complex_formulas),
    most_complex: Enum.take(complex_formulas, 5)
  }
end

# Example usage
analysis = analyze_spreadsheet_formulas("complex_financial_model.xlsx")
```

## Best Practices

1. **Clear Naming**: Use descriptive names for named ranges and defined names
2. **Avoid Hardcoding**: Use named ranges and defined names instead of hardcoded cell references
3. **Documentation**: Add comments in your code to explain what formulas are doing
4. **Error Handling**: Always check the return values of formula functions
5. **Cell References**: When referencing cells in other sheets, use the format `SheetName!CellAddress`

## Real-World Example: Financial Dashboard

Here's a comprehensive example that combines regular formulas, array formulas, named ranges, and defined names to create a financial dashboard:

```elixir
alias UmyaSpreadsheet

# Create a new spreadsheet for our financial dashboard
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Rename the default sheet to "Data"
sheet_names = UmyaSpreadsheet.get_sheet_names(spreadsheet)
first_sheet = List.first(sheet_names)

if first_sheet != "Data" do
  UmyaSpreadsheet.clone_sheet(spreadsheet, first_sheet, "Data")
  UmyaSpreadsheet.remove_sheet(spreadsheet, first_sheet)
end

# Add sheets for our dashboard
UmyaSpreadsheet.add_sheet(spreadsheet, "Dashboard")
UmyaSpreadsheet.add_sheet(spreadsheet, "Charts")

# Set Data sheet as active when opening
sheet_names = UmyaSpreadsheet.get_sheet_names(spreadsheet)
dashboard_index = Enum.find_index(sheet_names, fn name -> name == "Dashboard" end)
UmyaSpreadsheet.set_active_tab(spreadsheet, dashboard_index)

# ----- Data Sheet Setup -----

# Add monthly sales data for different regions
months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
regions = ["North", "South", "East", "West"]

# Create headers
UmyaSpreadsheet.set_cell_value(spreadsheet, "Data", "A1", "Month")
regions |> Enum.with_index(fn region, idx ->
  UmyaSpreadsheet.set_cell_value(spreadsheet, "Data", "#{<<66 + idx::utf8>>}1", region)
end)

# Add data rows
months |> Enum.with_index(fn month, row_idx ->
  row = row_idx + 2
  UmyaSpreadsheet.set_cell_value(spreadsheet, "Data", "A#{row}", month)

  regions |> Enum.with_index(fn _region, col_idx ->
    col = <<66 + col_idx::utf8>>
    # Generate some sample sales data (in a real app, this would be actual data)
    sales = :rand.uniform(10000) + 5000
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Data", "#{col}#{row}", sales)
  end)
end)

# Add cost percentages
UmyaSpreadsheet.set_cell_value(spreadsheet, "Data", "G1", "Cost Factors")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Data", "G2", "COGS %")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Data", "G3", "Marketing %")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Data", "G4", "Admin %")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Data", "H2", 0.38)  # 38% cost of goods sold
UmyaSpreadsheet.set_cell_value(spreadsheet, "Data", "H3", 0.15)  # 15% marketing expenses
UmyaSpreadsheet.set_cell_value(spreadsheet, "Data", "H4", 0.12)  # 12% administrative expenses

# Create named ranges for our data
UmyaSpreadsheet.create_named_range(spreadsheet, "MonthNames", "Data", "A2:A7")
UmyaSpreadsheet.create_named_range(spreadsheet, "NorthSales", "Data", "B2:B7")
UmyaSpreadsheet.create_named_range(spreadsheet, "SouthSales", "Data", "C2:C7")
UmyaSpreadsheet.create_named_range(spreadsheet, "EastSales", "Data", "D2:D7")
UmyaSpreadsheet.create_named_range(spreadsheet, "WestSales", "Data", "E2:E7")
UmyaSpreadsheet.create_named_range(spreadsheet, "AllSales", "Data", "B2:E7")

# Create defined names for our constants
UmyaSpreadsheet.create_defined_name(spreadsheet, "COGS_RATE", "Data!H2")
UmyaSpreadsheet.create_defined_name(spreadsheet, "MARKETING_RATE", "Data!H3")
UmyaSpreadsheet.create_defined_name(spreadsheet, "ADMIN_RATE", "Data!H4")
UmyaSpreadsheet.create_defined_name(spreadsheet, "PROFIT_RATE", "1-COGS_RATE-MARKETING_RATE-ADMIN_RATE")

# ----- Dashboard Sheet Setup -----

# Create a dashboard title
UmyaSpreadsheet.set_cell_value(spreadsheet, "Dashboard", "A1", "SALES DASHBOARD")
UmyaSpreadsheet.set_font_bold(spreadsheet, "Dashboard", "A1", true)
UmyaSpreadsheet.set_font_size(spreadsheet, "Dashboard", "A1", 16)

# Setup monthly totals section
UmyaSpreadsheet.set_cell_value(spreadsheet, "Dashboard", "A3", "Monthly Summary")
UmyaSpreadsheet.set_font_bold(spreadsheet, "Dashboard", "A3", true)

# Add month headers in the dashboard
UmyaSpreadsheet.set_cell_value(spreadsheet, "Dashboard", "A4", "Month")
months |> Enum.with_index(fn month, idx ->
  UmyaSpreadsheet.set_cell_value(spreadsheet, "Dashboard", "A#{5 + idx}", month)
end)

# Add KPI columns
metrics = ["Total Sales", "COGS", "Marketing", "Admin", "Profit"]
metrics |> Enum.with_index(fn metric, idx ->
  UmyaSpreadsheet.set_cell_value(spreadsheet, "Dashboard", "#{<<66 + idx::utf8>>}4", metric)
end)

# Create formulas for the dashboard calculations
months |> Enum.with_index(fn _month, row_idx ->
  row = row_idx + 5
  data_row = row_idx + 2  # Corresponding row in Data sheet

  # Total Sales (sum across all regions)
  UmyaSpreadsheet.set_formula(
    spreadsheet,
    "Dashboard",
    "B#{row}",
    "SUM(Data!B#{data_row}:E#{data_row})"
  )

  # COGS (using defined name)
  UmyaSpreadsheet.set_formula(
    spreadsheet,
    "Dashboard",
    "C#{row}",
    "B#{row}*COGS_RATE"
  )

  # Marketing expenses
  UmyaSpreadsheet.set_formula(
    spreadsheet,
    "Dashboard",
    "D#{row}",
    "B#{row}*MARKETING_RATE"
  )

  # Admin expenses
  UmyaSpreadsheet.set_formula(
    spreadsheet,
    "Dashboard",
    "E#{row}",
    "B#{row}*ADMIN_RATE"
  )

  # Profit calculation
  UmyaSpreadsheet.set_formula(
    spreadsheet,
    "Dashboard",
    "F#{row}",
    "B#{row}-C#{row}-D#{row}-E#{row}"
  )
end)

# Add totals row using array formulas
UmyaSpreadsheet.set_cell_value(spreadsheet, "Dashboard", "A11", "TOTAL")
UmyaSpreadsheet.set_font_bold(spreadsheet, "Dashboard", "A11", true)

["B", "C", "D", "E", "F"] |> Enum.each(fn col ->
  UmyaSpreadsheet.set_formula(
    spreadsheet,
    "Dashboard",
    "#{col}11",
    "SUM(#{col}5:#{col}10)"
  )
  UmyaSpreadsheet.set_font_bold(spreadsheet, "Dashboard", "#{col}11", true)
end)

# Add regional analysis section
UmyaSpreadsheet.set_cell_value(spreadsheet, "Dashboard", "A13", "Regional Analysis")
UmyaSpreadsheet.set_font_bold(spreadsheet, "Dashboard", "A13", true)

# Add region headers
UmyaSpreadsheet.set_cell_value(spreadsheet, "Dashboard", "A14", "Region")
regions |> Enum.with_index(fn region, idx ->
  UmyaSpreadsheet.set_cell_value(spreadsheet, "Dashboard", "A#{15 + idx}", region)
end)

# Add KPI columns for regions
["Total", "Average", "Min", "Max", "% of Total"] |> Enum.with_index(fn metric, idx ->
  UmyaSpreadsheet.set_cell_value(spreadsheet, "Dashboard", "#{<<66 + idx::utf8>>}14", metric)
end)

# Create array formulas for regional analysis
named_ranges = ["NorthSales", "SouthSales", "EastSales", "WestSales"]
named_ranges |> Enum.with_index(fn range_name, idx ->
  row = 15 + idx

  # Total for region
  UmyaSpreadsheet.set_formula(
    spreadsheet,
    "Dashboard",
    "B#{row}",
    "SUM(#{range_name})"
  )

  # Average for region
  UmyaSpreadsheet.set_formula(
    spreadsheet,
    "Dashboard",
    "C#{row}",
    "AVERAGE(#{range_name})"
  )

  # Min for region
  UmyaSpreadsheet.set_formula(
    spreadsheet,
    "Dashboard",
    "D#{row}",
    "MIN(#{range_name})"
  )

  # Max for region
  UmyaSpreadsheet.set_formula(
    spreadsheet,
    "Dashboard",
    "E#{row}",
    "MAX(#{range_name})"
  )

  # Percentage of total
  UmyaSpreadsheet.set_formula(
    spreadsheet,
    "Dashboard",
    "F#{row}",
    "B#{row}/B19"
  )
end)

# Add Grand Total for regions
UmyaSpreadsheet.set_cell_value(spreadsheet, "Dashboard", "A19", "GRAND TOTAL")
UmyaSpreadsheet.set_font_bold(spreadsheet, "Dashboard", "A19", true)

["B", "C", "D", "E"] |> Enum.each(fn col ->
  UmyaSpreadsheet.set_formula(
    spreadsheet,
    "Dashboard",
    "#{col}19",
    "SUM(#{col}15:#{col}18)"
  )
  UmyaSpreadsheet.set_font_bold(spreadsheet, "Dashboard", "#{col}19", true)
end)

# Set F19 to 100%
UmyaSpreadsheet.set_cell_value(spreadsheet, "Dashboard", "F19", 1)
UmyaSpreadsheet.set_font_bold(spreadsheet, "Dashboard", "F19", true)

# Format the dashboard
# Set number formatting
UmyaSpreadsheet.set_number_format(spreadsheet, "Dashboard", "B5:B19", "$#,##0")
UmyaSpreadsheet.set_number_format(spreadsheet, "Dashboard", "C5:F11", "$#,##0")
UmyaSpreadsheet.set_number_format(spreadsheet, "Dashboard", "B15:E19", "$#,##0")
UmyaSpreadsheet.set_number_format(spreadsheet, "Dashboard", "F15:F19", "0.0%")

# Add a chart to the Charts sheet
UmyaSpreadsheet.add_chart(
  spreadsheet,
  "Charts",
  "ColumnChart",
  "B2",
  "H15",
  "Monthly Sales by Region",
  ["Data!B2:B7", "Data!C2:C7", "Data!D2:D7", "Data!E2:E7"],
  "Data!A2:A7"
)

UmyaSpreadsheet.add_chart(
  spreadsheet,
  "Charts",
  "PieChart",
  "B17",
  "H30",
  "Regional Sales Distribution",
  ["Dashboard!B15:B18"],
  "Dashboard!A15:A18"
)

# Save the dashboard spreadsheet
UmyaSpreadsheet.write(spreadsheet, "sales_dashboard.xlsx")
```

## Limitations

- The library doesn't calculate formula results; it only stores the formula text
- Excel evaluates the formulas when the file is opened
- Some advanced Excel formula features may not be fully supported
- Array formulas require Excel to correctly interpret them
- Special care must be taken to format formulas correctly for Excel compatibility
