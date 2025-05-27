# Page Breaks

This guide covers how to work with page breaks in UmyaSpreadsheet. Page breaks control where pages split when printing or viewing spreadsheets in Page Break Preview mode.

## Table of Contents

- [Overview](#overview)
- [Types of Page Breaks](#types-of-page-breaks)
- [Basic Operations](#basic-operations)
- [Advanced Usage](#advanced-usage)
- [Best Practices](#best-practices)
- [Troubleshooting](#troubleshooting)

## Overview

Page breaks in Excel determine where content splits across pages when printing or viewing in Page Break Preview mode. UmyaSpreadsheet provides comprehensive page break management through the `UmyaSpreadsheet.PageBreaks` module.

### Key Features

- **Manual Page Breaks**: User-defined breaks that remain fixed
- **Row Page Breaks**: Horizontal breaks that split pages vertically
- **Column Page Breaks**: Vertical breaks that split pages horizontally
- **Bulk Operations**: Add multiple page breaks efficiently
- **Query Functions**: Check existing page breaks and their properties

## Types of Page Breaks

### Manual vs Automatic Page Breaks

Excel supports two types of page breaks:

- **Manual Page Breaks**: Explicitly set by users and have priority over automatic breaks
- **Automatic Page Breaks**: Calculated by Excel based on page size, margins, and content

UmyaSpreadsheet focuses on manual page breaks as they provide precise control over document layout.

### Row Page Breaks

Row page breaks create horizontal lines where pages split. Content above the break appears on one page, and content from the specified row onwards appears on the next page.

```elixir
# Add a page break above row 25
:ok = UmyaSpreadsheet.PageBreaks.add_row_page_break(spreadsheet, "Sheet1", 25)

# This means:
# - Rows 1-24 will be on the first page
# - Rows 25+ will be on subsequent pages
```

### Column Page Breaks

Column page breaks create vertical lines where pages split. Content to the left of the break appears on one page, and content from the specified column onwards appears on the next page.

```elixir
# Add a page break to the left of column 10 (column J)
:ok = UmyaSpreadsheet.PageBreaks.add_column_page_break(spreadsheet, "Sheet1", 10)

# This means:
# - Columns A-I will be on the left page
# - Columns J+ will be on subsequent pages
```

## Basic Operations

### Adding Page Breaks

#### Single Row Page Break

```elixir
# Add a manual page break above row 50
:ok = UmyaSpreadsheet.PageBreaks.add_row_page_break(spreadsheet, "Sales Report", 50)

# Add an automatic page break above row 75
:ok = UmyaSpreadsheet.PageBreaks.add_row_page_break(spreadsheet, "Sales Report", 75, false)
```

#### Single Column Page Break

```elixir
# Add a manual page break to the left of column 8 (column H)
:ok = UmyaSpreadsheet.PageBreaks.add_column_page_break(spreadsheet, "Sales Report", 8)

# Add an automatic page break to the left of column 12 (column L)
:ok = UmyaSpreadsheet.PageBreaks.add_column_page_break(spreadsheet, "Sales Report", 12, false)
```

### Removing Page Breaks

#### Remove Specific Page Breaks

```elixir
# Remove the page break above row 50
:ok = UmyaSpreadsheet.PageBreaks.remove_row_page_break(spreadsheet, "Sales Report", 50)

# Remove the page break to the left of column 8
:ok = UmyaSpreadsheet.PageBreaks.remove_column_page_break(spreadsheet, "Sales Report", 8)
```

#### Clear All Page Breaks

```elixir
# Clear all row page breaks
:ok = UmyaSpreadsheet.PageBreaks.clear_row_page_breaks(spreadsheet, "Sales Report")

# Clear all column page breaks
:ok = UmyaSpreadsheet.PageBreaks.clear_column_page_breaks(spreadsheet, "Sales Report")
```

### Querying Page Breaks

#### Check for Specific Page Breaks

```elixir
# Check if there's a page break above row 50
{:ok, has_break} = UmyaSpreadsheet.PageBreaks.has_row_page_break(spreadsheet, "Sales Report", 50)

if has_break do
  IO.puts("Page break found at row 50")
else
  IO.puts("No page break at row 50")
end

# Check if there's a page break to the left of column 8
{:ok, has_break} = UmyaSpreadsheet.PageBreaks.has_column_page_break(spreadsheet, "Sales Report", 8)
```

#### Get All Page Breaks

```elixir
# Get all row page breaks
{:ok, row_breaks} = UmyaSpreadsheet.PageBreaks.get_row_page_breaks(spreadsheet, "Sales Report")
# Returns: [{25, true}, {50, false}, {75, true}]
# Format: [{row_number, is_manual}, ...]

# Get all column page breaks
{:ok, column_breaks} = UmyaSpreadsheet.PageBreaks.get_column_page_breaks(spreadsheet, "Sales Report")
# Returns: [{8, true}, {15, true}]
# Format: [{column_number, is_manual}, ...]
```

#### Count Page Breaks

```elixir
# Count row page breaks
{:ok, row_count} = UmyaSpreadsheet.PageBreaks.count_row_page_breaks(spreadsheet, "Sales Report")
IO.puts("Total row page breaks: #{row_count}")

# Count column page breaks
{:ok, col_count} = UmyaSpreadsheet.PageBreaks.count_column_page_breaks(spreadsheet, "Sales Report")
IO.puts("Total column page breaks: #{col_count}")
```

## Advanced Usage

### Bulk Operations

For efficiency when adding multiple page breaks, use the bulk operation functions:

```elixir
# Add multiple row page breaks at once
:ok = UmyaSpreadsheet.PageBreaks.add_row_page_breaks(
  spreadsheet,
  "Financial Report",
  [25, 50, 75, 100, 125]
)

# Add multiple column page breaks at once
:ok = UmyaSpreadsheet.PageBreaks.add_column_page_breaks(
  spreadsheet,
  "Financial Report",
  [5, 10, 15, 20]
)

# Remove multiple page breaks efficiently
:ok = UmyaSpreadsheet.PageBreaks.remove_row_page_breaks(
  spreadsheet,
  "Financial Report",
  [25, 75]  # Remove some breaks
)

# Get all page breaks at once
{:ok, %{row_breaks: rows, column_breaks: cols}} =
  UmyaSpreadsheet.PageBreaks.get_all_page_breaks(spreadsheet, "Financial Report")

# Clear all page breaks from a sheet
:ok = UmyaSpreadsheet.PageBreaks.clear_all_page_breaks(spreadsheet, "Financial Report")
```

### Creating Report Sections

Page breaks are particularly useful for creating structured reports with distinct sections:

```elixir
defmodule ReportGenerator do
  def create_quarterly_report(spreadsheet) do
    # Add headers and data...

    # Create page breaks between quarters
    quarterly_breaks = [
      30,   # End of Q1 section
      60,   # End of Q2 section
      90,   # End of Q3 section
      120   # End of Q4 section
    ]

    :ok = UmyaSpreadsheet.PageBreaks.add_row_page_breaks(
      spreadsheet,
      "Quarterly Report",
      quarterly_breaks
    )

    # Add column breaks to separate different metric groups
    metric_breaks = [8, 15, 22]  # Separate revenue, costs, profit sections

    :ok = UmyaSpreadsheet.PageBreaks.add_column_page_breaks(
      spreadsheet,
      "Quarterly Report",
      metric_breaks
    )
  end
end
```

### Conditional Page Break Management

```elixir
defmodule PageBreakManager do
  def setup_dynamic_breaks(spreadsheet, sheet_name, data_rows) do
    # Only add page breaks if we have substantial data
    if data_rows > 50 do
      # Add page breaks every 25 rows for long reports
      breaks = Enum.take_every(25..data_rows, 25)

      UmyaSpreadsheet.PageBreaks.add_row_page_breaks(
        spreadsheet,
        sheet_name,
        breaks
      )
    end
  end

  def reset_page_breaks(spreadsheet, sheet_name) do
    # Clear existing breaks before applying new layout
    with :ok <- UmyaSpreadsheet.PageBreaks.clear_row_page_breaks(spreadsheet, sheet_name),
         :ok <- UmyaSpreadsheet.PageBreaks.clear_column_page_breaks(spreadsheet, sheet_name) do
      :ok
    end
  end
end
```

### Working with Page Break Information

```elixir
defmodule PageBreakAnalyzer do
  def analyze_page_breaks(spreadsheet, sheet_name) do
    # Get all page breaks
    {:ok, row_breaks} = UmyaSpreadsheet.PageBreaks.get_row_page_breaks(spreadsheet, sheet_name)
    {:ok, col_breaks} = UmyaSpreadsheet.PageBreaks.get_column_page_breaks(spreadsheet, sheet_name)

    # Analyze the data
    manual_row_breaks = Enum.filter(row_breaks, fn {_row, manual} -> manual end)
    auto_row_breaks = Enum.filter(row_breaks, fn {_row, manual} -> not manual end)

    manual_col_breaks = Enum.filter(col_breaks, fn {_col, manual} -> manual end)
    auto_col_breaks = Enum.filter(col_breaks, fn {_col, manual} -> not manual end)

    %{
      total_row_breaks: length(row_breaks),
      manual_row_breaks: length(manual_row_breaks),
      automatic_row_breaks: length(auto_row_breaks),
      total_column_breaks: length(col_breaks),
      manual_column_breaks: length(manual_col_breaks),
      automatic_column_breaks: length(auto_col_breaks),
      row_break_positions: Enum.map(row_breaks, fn {row, _} -> row end),
      column_break_positions: Enum.map(col_breaks, fn {col, _} -> col end)
    }
  end
end
```

## Best Practices

### 1. Strategic Placement

- **Place breaks at logical boundaries**: End of sections, after headers, before summaries
- **Consider content flow**: Ensure related data stays together
- **Account for print margins**: Leave sufficient space for headers and footers

### 2. Consistency

```elixir
# Use consistent spacing for similar reports
defmodule ReportStandards do
  @section_break_interval 25
  @subsection_break_interval 50

  def apply_standard_breaks(spreadsheet, sheet_name, total_rows) do
    # Apply breaks at standard intervals
    major_breaks = Enum.take_every(@subsection_break_interval..total_rows, @subsection_break_interval)

    UmyaSpreadsheet.PageBreaks.add_row_page_breaks(
      spreadsheet,
      sheet_name,
      major_breaks
    )
  end
end
```

### 3. Performance Optimization

```elixir
# Use bulk operations for better performance
# Instead of:
# Enum.each(row_numbers, fn row ->
#   UmyaSpreadsheet.PageBreaks.add_row_page_break(spreadsheet, sheet_name, row)
# end)

# Use:
UmyaSpreadsheet.PageBreaks.add_row_page_breaks(
  spreadsheet,
  sheet_name,
  row_numbers
)
```

### 4. Error Handling

```elixir
defmodule SafePageBreaks do
  def add_row_break_safely(spreadsheet, sheet_name, row_number) do
    case UmyaSpreadsheet.PageBreaks.add_row_page_break(spreadsheet, sheet_name, row_number) do
      :ok ->
        {:ok, "Page break added at row #{row_number}"}

      {:error, :sheet_not_found} ->
        {:error, "Sheet '#{sheet_name}' not found"}

      {:error, reason} ->
        {:error, "Failed to add page break: #{reason}"}
    end
  end
end
```

### 5. Testing Page Breaks

```elixir
defmodule PageBreakTests do
  def verify_page_breaks(spreadsheet, sheet_name, expected_breaks) do
    {:ok, actual_breaks} = UmyaSpreadsheet.PageBreaks.get_row_page_breaks(spreadsheet, sheet_name)
    actual_positions = Enum.map(actual_breaks, fn {row, _} -> row end)

    case Enum.sort(actual_positions) == Enum.sort(expected_breaks) do
      true ->
        {:ok, "All expected page breaks found"}

      false ->
        missing = expected_breaks -- actual_positions
        extra = actual_positions -- expected_breaks

        {:error, %{
          missing: missing,
          extra: extra,
          expected: expected_breaks,
          actual: actual_positions
        }}
    end
  end
end
```

## Real-World Examples

### Financial Report with Sections

```elixir
defmodule FinancialReport do
  def create_report(spreadsheet) do
    sheet_name = "Financial Report 2024"

    # Add the sheet
    UmyaSpreadsheet.add_sheet(spreadsheet, sheet_name)

    # Add headers and data (simplified)
    UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, "A1", "Financial Report 2024")

    # Define section boundaries
    sections = [
      {30, "Revenue Section"},
      {60, "Expenses Section"},
      {90, "Profit Analysis"},
      {120, "Year-over-Year Comparison"}
    ]

    # Add page breaks between sections
    section_breaks = Enum.map(sections, fn {row, _name} -> row end)

    UmyaSpreadsheet.PageBreaks.add_row_page_breaks(
      spreadsheet,
      sheet_name,
      section_breaks
    )

    # Add column breaks to separate quarters
    quarter_breaks = [8, 15, 22]  # After Q1, Q2, Q3

    UmyaSpreadsheet.PageBreaks.add_column_page_breaks(
      spreadsheet,
      sheet_name,
      quarter_breaks
    )

    IO.puts("Financial report created with structured page breaks")
  end
end
```

### Large Dataset with Regular Breaks

```elixir
defmodule LargeDatasetProcessor do
  def process_large_dataset(spreadsheet, data) do
    sheet_name = "Large Dataset"
    rows_per_page = 50

    # Calculate total rows
    total_rows = length(data)

    # Add data (simplified)
    # ... data insertion logic ...

    # Calculate page break positions
    page_breaks =
      rows_per_page
      |> Stream.iterate(&(&1 + rows_per_page))
      |> Stream.take_while(&(&1 < total_rows))
      |> Enum.to_list()

    # Add page breaks efficiently
    case UmyaSpreadsheet.PageBreaks.add_row_page_breaks(
      spreadsheet,
      sheet_name,
      page_breaks
    ) do
      :ok ->
        IO.puts("Added #{length(page_breaks)} page breaks for #{total_rows} rows")

      {:error, reason} ->
        IO.puts("Failed to add page breaks: #{reason}")
    end
  end
end
```

## Troubleshooting

### Common Issues

#### Issue: Page breaks not visible in Excel

**Cause**: Page breaks are only visible in Page Break Preview mode or when printing.

**Solution**:

- Switch to Page Break Preview mode in Excel (View > Page Break Preview)
- Check print preview to see page breaks in action

#### Issue: "Sheet not found" errors

**Cause**: Attempting to add page breaks to a non-existent sheet.

**Solution**:

```elixir
# Always verify sheet exists before adding page breaks
case UmyaSpreadsheet.get_sheet_names(spreadsheet) do
  {:ok, sheet_names} ->
    if sheet_name in sheet_names do
      UmyaSpreadsheet.PageBreaks.add_row_page_break(spreadsheet, sheet_name, row)
    else
      {:error, "Sheet '#{sheet_name}' does not exist"}
    end

  error -> error
end
```

#### Issue: Unexpected page break behavior

**Cause**: Mixing manual and automatic page breaks can cause confusion.

**Solution**: Use consistent approach - either all manual or let Excel handle automatic breaks.

### Debugging Page Breaks

```elixir
defmodule PageBreakDebugger do
  def debug_page_breaks(spreadsheet, sheet_name) do
    IO.puts("=== Page Break Debug Info for '#{sheet_name}' ===")

    # Get and display row breaks
    case UmyaSpreadsheet.PageBreaks.get_row_page_breaks(spreadsheet, sheet_name) do
      {:ok, row_breaks} ->
        IO.puts("Row Page Breaks (#{length(row_breaks)}):")
        Enum.each(row_breaks, fn {row, manual} ->
          type = if manual, do: "Manual", else: "Automatic"
          IO.puts("  Row #{row}: #{type}")
        end)

      {:error, reason} ->
        IO.puts("Error getting row breaks: #{reason}")
    end

    # Get and display column breaks
    case UmyaSpreadsheet.PageBreaks.get_column_page_breaks(spreadsheet, sheet_name) do
      {:ok, col_breaks} ->
        IO.puts("Column Page Breaks (#{length(col_breaks)}):")
        Enum.each(col_breaks, fn {col, manual} ->
          type = if manual, do: "Manual", else: "Automatic"
          IO.puts("  Column #{col}: #{type}")
        end)

      {:error, reason} ->
        IO.puts("Error getting column breaks: #{reason}")
    end

    IO.puts("=== End Debug Info ===")
  end
end
```

### Performance Considerations

- **Use bulk operations**: `add_row_page_breaks/3`, `add_column_page_breaks/3`, and other bulk functions are more efficient than individual calls
- **Clear existing breaks**: Use `clear_*` functions before applying new page break schemes
- **Minimize break modifications**: Adding/removing breaks frequently can impact performance
- **Test with realistic data**: Page break behavior can change with different data volumes

This guide provides comprehensive coverage of page break functionality in UmyaSpreadsheet. Page breaks are powerful tools for controlling document layout and ensuring professional presentation of your spreadsheet data.
