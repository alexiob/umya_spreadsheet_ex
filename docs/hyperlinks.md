# Excel Hyperlinks Guide

This guide covers the comprehensive hyperlink functionality available in UmyaSpreadsheet, including creating, managing, and removing hyperlinks from Excel spreadsheets.

## Table of Contents

1. [Overview](#overview)
2. [Quick Start](#quick-start)
3. [Basic Hyperlink Operations](#basic-hyperlink-operations)
4. [Advanced Hyperlink Features](#advanced-hyperlink-features)
5. [Hyperlink Types and Examples](#hyperlink-types-and-examples)
6. [Best Practices](#best-practices)
7. [Error Handling](#error-handling)
8. [Performance Considerations](#performance-considerations)
9. [Real-World Examples](#real-world-examples)

## Overview

UmyaSpreadsheet provides a complete API for managing hyperlinks in Excel files. Hyperlinks can point to:

- **Web URLs** (HTTP/HTTPS) - External websites and web services
- **Email addresses** (mailto) - Direct email composition
- **File paths** - Local and network files
- **Internal references** - Other worksheets and cell ranges within the same workbook

### Key Features

- ✅ **Complete CRUD Operations** - Create, Read, Update, Delete hyperlinks
- ✅ **Multiple Hyperlink Types** - Web, email, file, and internal references
- ✅ **Rich Metadata** - Custom tooltips and descriptions
- ✅ **Bulk Operations** - Add or remove multiple hyperlinks efficiently
- ✅ **Integration with Cell Values** - Hyperlinks work alongside cell content
- ✅ **Worksheet-Level Management** - Query and manage all hyperlinks in a sheet

## Quick Start

Here's a simple example to get you started with hyperlinks:

```elixir
# Create a new spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add a web URL hyperlink
{:ok, _} = UmyaSpreadsheet.add_hyperlink(
  spreadsheet,
  "Sheet1",
  "A1",
  "https://example.com",
  "Visit our website"
)

# Add an internal reference
{:ok, _} = UmyaSpreadsheet.add_hyperlink(
  spreadsheet,
  "Sheet1",
  "B1",
  "Sheet2!A1",
  "Go to Sheet2",
  true  # is_internal = true
)

# Get hyperlink information
{:ok, hyperlink_info} = UmyaSpreadsheet.get_hyperlink(spreadsheet, "Sheet1", "A1")
IO.inspect(hyperlink_info)
# Output: %{
#   "url" => "https://example.com",
#   "tooltip" => "Visit our website",
#   "location" => "A1"
# }

# Save the spreadsheet
:ok = UmyaSpreadsheet.write(spreadsheet, "output/hyperlinks_example.xlsx")
```

## Basic Hyperlink Operations

### Adding Hyperlinks

#### Simple Web URL

```elixir
# Basic hyperlink without tooltip
:ok = UmyaSpreadsheet.add_hyperlink(
  spreadsheet,
  "Sheet1",
  "A1",
  "https://google.com"
)

# Hyperlink with tooltip
:ok = UmyaSpreadsheet.add_hyperlink(
  spreadsheet,
  "Sheet1",
  "A2",
  "https://github.com",
  "GitHub Repository"
)
```

#### Email Links

```elixir
:ok = UmyaSpreadsheet.add_hyperlink(
  spreadsheet,
  "Sheet1",
  "B1",
  "mailto:contact@example.com",
  "Send us an email"
)

# Email with subject and body
:ok = UmyaSpreadsheet.add_hyperlink(
  spreadsheet,
  "Sheet1",
  "B2",
  "mailto:support@example.com?subject=Support%20Request&body=Hello,%20I%20need%20help",
  "Contact Support"
)
```

#### File Path Links

```elixir
# Local file
:ok = UmyaSpreadsheet.add_hyperlink(
  spreadsheet,
  "Sheet1",
  "C1",
  "file:///Users/username/Documents/report.pdf",
  "Open Report"
)

# Network file
:ok = UmyaSpreadsheet.add_hyperlink(
  spreadsheet,
  "Sheet1",
  "C2",
  "file://server/shared/documents/manual.docx",
  "User Manual"
)
```

#### Internal Worksheet References

```elixir
# Reference to another sheet
:ok = UmyaSpreadsheet.add_hyperlink(
  spreadsheet,
  "Sheet1",
  "D1",
  "Sheet2!A1",
  "Go to Sheet2 Cell A1",
  true  # is_internal = true
)

# Reference to a range
:ok = UmyaSpreadsheet.add_hyperlink(
  spreadsheet,
  "Sheet1",
  "D2",
  "Data!A1:C10",
  "View Data Range",
  true
)
```

### Getting Hyperlink Information

```elixir
# Get single hyperlink
case UmyaSpreadsheet.get_hyperlink(spreadsheet, "Sheet1", "A1") do
  {:ok, hyperlink_info} ->
    url = hyperlink_info["url"]
    tooltip = hyperlink_info["tooltip"]
    location = hyperlink_info["location"]
    IO.puts("Found hyperlink: #{url}")

  {:error, :not_found} ->
    IO.puts("No hyperlink found in cell A1")

  {:error, reason} ->
    IO.puts("Error: #{reason}")
end
```

### Checking for Hyperlinks

```elixir
# Check if a specific cell has a hyperlink
{:ok, has_link} = UmyaSpreadsheet.has_hyperlink(spreadsheet, "Sheet1", "A1")

if has_link do
  IO.puts("Cell A1 has a hyperlink")
else
  IO.puts("Cell A1 does not have a hyperlink")
end

# Check if worksheet has any hyperlinks
{:ok, has_any} = UmyaSpreadsheet.has_hyperlinks(spreadsheet, "Sheet1")
```

### Updating Hyperlinks

```elixir
# Update existing hyperlink
:ok = UmyaSpreadsheet.update_hyperlink(
  spreadsheet,
  "Sheet1",
  "A1",
  "https://newexample.com",
  "Updated tooltip"
)

# Convert external to internal reference
:ok = UmyaSpreadsheet.update_hyperlink(
  spreadsheet,
  "Sheet1",
  "A1",
  "Summary!B5",
  "Go to Summary",
  true  # is_internal = true
)
```

### Removing Hyperlinks

```elixir
# Remove single hyperlink (preserves cell value)
:ok = UmyaSpreadsheet.remove_hyperlink(spreadsheet, "Sheet1", "A1")

# Remove all hyperlinks from worksheet
:ok = UmyaSpreadsheet.remove_all_hyperlinks(spreadsheet, "Sheet1")
```

## Advanced Hyperlink Features

### Getting All Hyperlinks from a Worksheet

```elixir
{:ok, hyperlinks} = UmyaSpreadsheet.get_hyperlinks(spreadsheet, "Sheet1")

Enum.each(hyperlinks, fn hyperlink ->
  IO.puts("Cell: #{hyperlink["cell"]}")
  IO.puts("URL: #{hyperlink["url"]}")
  IO.puts("Tooltip: #{hyperlink["tooltip"]}")
  IO.puts("---")
end)
```

### Bulk Hyperlink Operations

```elixir
# Add multiple hyperlinks at once
hyperlinks = [
  {"A1", "https://example.com", "Website"},
  {"A2", "mailto:test@example.com", "Email"},
  {"A3", "Sheet2!A1", "Internal link", true},
  {"A4", "file:///path/to/file.pdf", "Document"}
]

:ok = UmyaSpreadsheet.add_bulk_hyperlinks(spreadsheet, "Sheet1", hyperlinks)

# Using map format
hyperlinks_map = [
  %{cell: "B1", url: "https://github.com", tooltip: "GitHub"},
  %{cell: "B2", url: "mailto:dev@example.com", tooltip: "Developer Email"},
  %{cell: "B3", url: "Config!A1", tooltip: "Config Sheet", is_internal: true}
]

:ok = UmyaSpreadsheet.add_bulk_hyperlinks(spreadsheet, "Sheet1", hyperlinks_map)
```

### Counting Hyperlinks

```elixir
{:ok, count} = UmyaSpreadsheet.count_hyperlinks(spreadsheet, "Sheet1")
IO.puts("Sheet1 has #{count} hyperlinks")
```

### Hyperlinks with Cell Values

```elixir
# Set cell value and add hyperlink
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Click Here")
:ok = UmyaSpreadsheet.add_hyperlink(
  spreadsheet,
  "Sheet1",
  "A1",
  "https://example.com",
  "Visit our website"
)

# The cell will display "Click Here" as clickable text
```

## Hyperlink Types and Examples

### 1. Web URLs

```elixir
# HTTPS (secure)
:ok = UmyaSpreadsheet.add_hyperlink(
  spreadsheet, "Sheet1", "A1",
  "https://secure.example.com", "Secure Site"
)

# HTTP (unsecure)
:ok = UmyaSpreadsheet.add_hyperlink(
  spreadsheet, "Sheet1", "A2",
  "http://example.com", "Regular Site"
)

# FTP
:ok = UmyaSpreadsheet.add_hyperlink(
  spreadsheet, "Sheet1", "A3",
  "ftp://files.example.com", "FTP Server"
)
```

### 2. Email Addresses

```elixir
# Simple email
:ok = UmyaSpreadsheet.add_hyperlink(
  spreadsheet, "Sheet1", "B1",
  "mailto:contact@example.com", "Contact Us"
)

# Email with subject
:ok = UmyaSpreadsheet.add_hyperlink(
  spreadsheet, "Sheet1", "B2",
  "mailto:support@example.com?subject=Help%20Request", "Get Help"
)

# Email with multiple recipients
:ok = UmyaSpreadsheet.add_hyperlink(
  spreadsheet, "Sheet1", "B3",
  "mailto:sales@example.com?cc=manager@example.com&bcc=archive@example.com",
  "Contact Sales Team"
)
```

### 3. File Paths

```elixir
# Windows local file
:ok = UmyaSpreadsheet.add_hyperlink(
  spreadsheet, "Sheet1", "C1",
  "file:///C:/Documents/report.pdf", "Windows Report"
)

# macOS/Linux local file
:ok = UmyaSpreadsheet.add_hyperlink(
  spreadsheet, "Sheet1", "C2",
  "file:///Users/username/Documents/data.xlsx", "Data File"
)

# Network share (Windows)
:ok = UmyaSpreadsheet.add_hyperlink(
  spreadsheet, "Sheet1", "C3",
  "file://server/share/documents/manual.docx", "Network Manual"
)
```

### 4. Internal References

```elixir
# Reference to specific cell
:ok = UmyaSpreadsheet.add_hyperlink(
  spreadsheet, "Sheet1", "D1",
  "Summary!B5", "Go to Summary B5", true
)

# Reference to named range
:ok = UmyaSpreadsheet.add_hyperlink(
  spreadsheet, "Sheet1", "D2",
  "DataTable", "View Data Table", true
)

# Reference to cell range
:ok = UmyaSpreadsheet.add_hyperlink(
  spreadsheet, "Sheet1", "D3",
  "Charts!A1:F20", "View Charts Area", true
)
```

## Best Practices

### 1. URL Validation

```elixir
defmodule HyperlinkHelpers do
  def validate_url(url) do
    case URI.parse(url) do
      %URI{scheme: scheme} when scheme in ["http", "https", "mailto", "file"] ->
        {:ok, url}
      _ ->
        {:error, :invalid_url}
    end
  end

  def add_safe_hyperlink(spreadsheet, sheet_name, cell, url, tooltip \\ nil) do
    case validate_url(url) do
      {:ok, valid_url} ->
        UmyaSpreadsheet.add_hyperlink(spreadsheet, sheet_name, cell, valid_url, tooltip)
      {:error, reason} ->
        {:error, reason}
    end
  end
end

# Usage
case HyperlinkHelpers.add_safe_hyperlink(spreadsheet, "Sheet1", "A1", "https://example.com") do
  :ok -> IO.puts("Hyperlink added successfully")
  {:error, :invalid_url} -> IO.puts("Invalid URL format")
end
```

### 2. Tooltips for User Experience

```elixir
# Descriptive tooltips help users understand the link destination
:ok = UmyaSpreadsheet.add_hyperlink(
  spreadsheet, "Sheet1", "A1",
  "https://api.example.com/docs",
  "API Documentation - Authentication, endpoints, and examples"
)

:ok = UmyaSpreadsheet.add_hyperlink(
  spreadsheet, "Sheet1", "A2",
  "mailto:support@example.com",
  "Contact technical support (Response within 24 hours)"
)
```

### 3. Organizing Hyperlinks

```elixir
defmodule SpreadsheetBuilder do
  def add_navigation_section(spreadsheet, sheet_name) do
    navigation_links = [
      {"A1", "Summary!A1", "📊 Summary Dashboard", true},
      {"A2", "Data!A1", "📈 Raw Data", true},
      {"A3", "Charts!A1", "📉 Charts & Graphs", true},
      {"A4", "Settings!A1", "⚙️ Configuration", true}
    ]

    # Add section header
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, "A0", "Navigation")

    # Add all navigation links
    Enum.each(navigation_links, fn {cell, url, tooltip, is_internal} ->
      UmyaSpreadsheet.add_hyperlink(spreadsheet, sheet_name, cell, url, tooltip, is_internal)
    end)
  end
end
```

### 4. Error Recovery

```elixir
defmodule HyperlinkManager do
  def add_with_retry(spreadsheet, sheet_name, cell, url, tooltip, retries \\ 3) do
    case UmyaSpreadsheet.add_hyperlink(spreadsheet, sheet_name, cell, url, tooltip) do
      :ok ->
        :ok
      {:error, reason} when retries > 0 ->
        :timer.sleep(100)
        add_with_retry(spreadsheet, sheet_name, cell, url, tooltip, retries - 1)
      {:error, reason} ->
        {:error, reason}
    end
  end

  def safe_bulk_add(spreadsheet, sheet_name, hyperlinks) do
    Enum.reduce_while(hyperlinks, [], fn hyperlink_spec, acc ->
      case process_hyperlink(spreadsheet, sheet_name, hyperlink_spec) do
        :ok ->
          {:cont, [hyperlink_spec | acc]}
        {:error, reason} ->
          {:halt, {:error, reason, acc}}
      end
    end)
  end

  defp process_hyperlink(spreadsheet, sheet_name, {cell, url, tooltip, is_internal}) do
    UmyaSpreadsheet.add_hyperlink(spreadsheet, sheet_name, cell, url, tooltip, is_internal)
  end

  defp process_hyperlink(spreadsheet, sheet_name, {cell, url, tooltip}) do
    UmyaSpreadsheet.add_hyperlink(spreadsheet, sheet_name, cell, url, tooltip)
  end

  defp process_hyperlink(spreadsheet, sheet_name, {cell, url}) do
    UmyaSpreadsheet.add_hyperlink(spreadsheet, sheet_name, cell, url)
  end
end
```

## Error Handling

### Common Error Scenarios

```elixir
# Handle missing sheets
case UmyaSpreadsheet.add_hyperlink(spreadsheet, "NonExistentSheet", "A1", "https://example.com") do
  :ok ->
    IO.puts("Success")
  {:error, "Sheet not found"} ->
    IO.puts("Sheet doesn't exist - creating it first")
    :ok = UmyaSpreadsheet.add_worksheet(spreadsheet, "NonExistentSheet")
    :ok = UmyaSpreadsheet.add_hyperlink(spreadsheet, "NonExistentSheet", "A1", "https://example.com")
end

# Handle invalid cell references
case UmyaSpreadsheet.add_hyperlink(spreadsheet, "Sheet1", "INVALID", "https://example.com") do
  :ok ->
    IO.puts("Success")
  {:error, reason} ->
    IO.puts("Invalid cell reference: #{reason}")
end

# Handle update of non-existent hyperlinks
case UmyaSpreadsheet.update_hyperlink(spreadsheet, "Sheet1", "A1", "https://new.com") do
  :ok ->
    IO.puts("Updated successfully")
  {:error, "No hyperlink found in cell to update"} ->
    IO.puts("No existing hyperlink - adding new one")
    :ok = UmyaSpreadsheet.add_hyperlink(spreadsheet, "Sheet1", "A1", "https://new.com")
end
```

### Error Recovery Patterns

```elixir
defmodule RobustHyperlinkOperations do
  def ensure_hyperlink_exists(spreadsheet, sheet_name, cell, url, tooltip \\ nil) do
    case UmyaSpreadsheet.get_hyperlink(spreadsheet, sheet_name, cell) do
      {:ok, _existing} ->
        # Hyperlink exists, optionally update it
        UmyaSpreadsheet.update_hyperlink(spreadsheet, sheet_name, cell, url, tooltip)

      {:error, :not_found} ->
        # No hyperlink, add new one
        UmyaSpreadsheet.add_hyperlink(spreadsheet, sheet_name, cell, url, tooltip)

      {:error, reason} ->
        {:error, reason}
    end
  end

  def cleanup_broken_hyperlinks(spreadsheet, sheet_name) do
    case UmyaSpreadsheet.get_hyperlinks(spreadsheet, sheet_name) do
      {:ok, hyperlinks} ->
        broken_links = Enum.filter(hyperlinks, &is_broken_link?/1)

        Enum.each(broken_links, fn %{"cell" => cell} ->
          UmyaSpreadsheet.remove_hyperlink(spreadsheet, sheet_name, cell)
        end)

        {:ok, length(broken_links)}

      {:error, reason} ->
        {:error, reason}
    end
  end

  defp is_broken_link?(%{"url" => url}) do
    # Custom logic to check if URL is broken
    String.contains?(url, "broken") or String.length(url) == 0
  end
end
```

## Performance Considerations

### 1. Bulk Operations

```elixir
# ❌ Inefficient - Multiple individual calls
Enum.each(1..100, fn i ->
  UmyaSpreadsheet.add_hyperlink(
    spreadsheet, "Sheet1", "A#{i}",
    "https://example.com/page#{i}",
    "Page #{i}"
  )
end)

# ✅ Efficient - Single bulk operation
hyperlinks = Enum.map(1..100, fn i ->
  {"A#{i}", "https://example.com/page#{i}", "Page #{i}"}
end)

:ok = UmyaSpreadsheet.add_bulk_hyperlinks(spreadsheet, "Sheet1", hyperlinks)
```

### 2. Memory Management

```elixir
# Process hyperlinks in batches for large datasets
defmodule BatchProcessor do
  def process_hyperlinks_in_batches(spreadsheet, sheet_name, hyperlinks, batch_size \\ 50) do
    hyperlinks
    |> Enum.chunk_every(batch_size)
    |> Enum.reduce(:ok, fn batch, :ok ->
      UmyaSpreadsheet.add_bulk_hyperlinks(spreadsheet, sheet_name, batch)
    end)
  end
end
```

### 3. Lazy Loading

```elixir
# Only load hyperlink information when needed
defmodule LazyHyperlinkInfo do
  def get_hyperlink_lazy(spreadsheet, sheet_name, cell) do
    fn ->
      UmyaSpreadsheet.get_hyperlink(spreadsheet, sheet_name, cell)
    end
  end

  def get_all_hyperlinks_stream(spreadsheet, sheet_name) do
    case UmyaSpreadsheet.get_hyperlinks(spreadsheet, sheet_name) do
      {:ok, hyperlinks} ->
        Stream.map(hyperlinks, fn hyperlink ->
          # Process each hyperlink lazily
          process_hyperlink_info(hyperlink)
        end)
      {:error, reason} ->
        Stream.repeatedly(fn -> {:error, reason} end) |> Stream.take(1)
    end
  end

  defp process_hyperlink_info(hyperlink) do
    # Expensive processing only when accessed
    hyperlink
  end
end
```

## Real-World Examples

### Example 1: Employee Directory with Contact Links

```elixir
defmodule EmployeeDirectory do
  def create_directory(employees) do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add headers
    headers = ["Name", "Email", "Department", "Profile"]
    Enum.with_index(headers, fn header, index ->
      column = get_column_letter(index)
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "#{column}1", header)
    end)

    # Add employee data with hyperlinks
    Enum.with_index(employees, 2, fn employee, row ->
      # Name
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A#{row}", employee.name)

      # Email with mailto link
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B#{row}", employee.email)
      UmyaSpreadsheet.add_hyperlink(
        spreadsheet, "Sheet1", "B#{row}",
        "mailto:#{employee.email}",
        "Send email to #{employee.name}"
      )

      # Department
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C#{row}", employee.department)

      # Profile link
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D#{row}", "View Profile")
      UmyaSpreadsheet.add_hyperlink(
        spreadsheet, "Sheet1", "D#{row}",
        "https://company.com/profiles/#{employee.id}",
        "View #{employee.name}'s profile"
      )
    end)

    UmyaSpreadsheet.write(spreadsheet, "employee_directory.xlsx")
  end

  defp get_column_letter(index) do
    # Convert 0,1,2,3 to A,B,C,D
    <<65 + index::utf8>>
  end
end

# Usage
employees = [
  %{name: "John Doe", email: "john@company.com", department: "Engineering", id: "emp001"},
  %{name: "Jane Smith", email: "jane@company.com", department: "Marketing", id: "emp002"},
  %{name: "Bob Johnson", email: "bob@company.com", department: "Sales", id: "emp003"}
]

EmployeeDirectory.create_directory(employees)
```

### Example 2: Financial Report with Navigation

```elixir
defmodule FinancialReport do
  def create_report(financial_data) do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Create multiple sheets
    sheets = ["Summary", "Income", "Expenses", "Balance"]
    Enum.each(sheets, fn sheet ->
      if sheet != "Sheet1" do
        UmyaSpreadsheet.add_worksheet(spreadsheet, sheet)
      else
        UmyaSpreadsheet.rename_worksheet(spreadsheet, "Sheet1", sheet)
      end
    end)

    # Add navigation in Summary sheet
    add_navigation(spreadsheet, "Summary")

    # Add external links for regulatory information
    add_regulatory_links(spreadsheet, "Summary")

    UmyaSpreadsheet.write(spreadsheet, "financial_report.xlsx")
  end

  defp add_navigation(spreadsheet, sheet_name) do
    # Navigation header
    UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, "A1", "Quick Navigation")

    # Internal navigation links
    navigation_links = [
      {"A3", "Income!A1", "📊 Income Statement", true},
      {"A4", "Expenses!A1", "💰 Expense Report", true},
      {"A5", "Balance!A1", "⚖️ Balance Sheet", true}
    ]

    Enum.each(navigation_links, fn {cell, url, tooltip, is_internal} ->
      # Set display text
      display_text = tooltip |> String.replace(~r/[📊💰⚖️] /, "")
      UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, cell, display_text)

      # Add hyperlink
      UmyaSpreadsheet.add_hyperlink(spreadsheet, sheet_name, cell, url, tooltip, is_internal)
    end)
  end

  defp add_regulatory_links(spreadsheet, sheet_name) do
    # Regulatory information header
    UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, "A7", "Regulatory Information")

    # External regulatory links
    regulatory_links = [
      {"A9", "https://www.sec.gov", "SEC - Securities and Exchange Commission"},
      {"A10", "https://www.fasb.org", "FASB - Financial Accounting Standards Board"},
      {"A11", "https://www.irs.gov", "IRS - Internal Revenue Service"}
    ]

    Enum.each(regulatory_links, fn {cell, url, tooltip} ->
      # Set display text
      display_text = tooltip |> String.split(" - ") |> List.first()
      UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, cell, display_text)

      # Add hyperlink
      UmyaSpreadsheet.add_hyperlink(spreadsheet, sheet_name, cell, url, tooltip)
    end)
  end
end

# Usage
financial_data = %{
  revenue: 1_000_000,
  expenses: 750_000,
  profit: 250_000
}

FinancialReport.create_report(financial_data)
```

### Example 3: Project Documentation Hub

```elixir
defmodule ProjectDocumentationHub do
  def create_hub(project_info) do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Project overview
    add_project_overview(spreadsheet, "Sheet1", project_info)

    # Documentation links
    add_documentation_links(spreadsheet, "Sheet1", project_info.docs)

    # Team contact information
    add_team_contacts(spreadsheet, "Sheet1", project_info.team)

    # Resource links
    add_resource_links(spreadsheet, "Sheet1", project_info.resources)

    UmyaSpreadsheet.write(spreadsheet, "project_hub.xlsx")
  end

  defp add_project_overview(spreadsheet, sheet_name, project_info) do
    UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, "A1", "Project: #{project_info.name}")
    UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, "A2", "Status: #{project_info.status}")

    # Project repository link
    UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, "A3", "Repository")
    UmyaSpreadsheet.add_hyperlink(
      spreadsheet, sheet_name, "A3",
      project_info.repository_url,
      "View project source code on GitHub"
    )
  end

  defp add_documentation_links(spreadsheet, sheet_name, docs) do
    UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, "A5", "Documentation")

    Enum.with_index(docs, 6, fn doc, row ->
      UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, "A#{row}", doc.title)
      UmyaSpreadsheet.add_hyperlink(
        spreadsheet, sheet_name, "A#{row}",
        doc.url,
        doc.description
      )
    end)
  end

  defp add_team_contacts(spreadsheet, sheet_name, team) do
    start_row = 6 + length(team) + 2
    UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, "A#{start_row}", "Team Contacts")

    Enum.with_index(team, start_row + 1, fn member, row ->
      UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, "A#{row}", member.name)
      UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, "B#{row}", member.role)

      # Email link
      UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, "C#{row}", "Email")
      UmyaSpreadsheet.add_hyperlink(
        spreadsheet, sheet_name, "C#{row}",
        "mailto:#{member.email}",
        "Send email to #{member.name}"
      )

      # Slack link (if available)
      if member.slack_id do
        UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, "D#{row}", "Slack")
        UmyaSpreadsheet.add_hyperlink(
          spreadsheet, sheet_name, "D#{row}",
          "slack://user?team=#{member.team_id}&id=#{member.slack_id}",
          "Message #{member.name} on Slack"
        )
      end
    end)
  end

  defp add_resource_links(spreadsheet, sheet_name, resources) do
    start_row = 20  # Fixed position for resources
    UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, "A#{start_row}", "Resources")

    Enum.with_index(resources, start_row + 1, fn resource, row ->
      UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, "A#{row}", resource.name)
      UmyaSpreadsheet.add_hyperlink(
        spreadsheet, sheet_name, "A#{row}",
        resource.url,
        resource.description
      )
    end)
  end
end

# Usage
project_info = %{
  name: "UmyaSpreadsheet",
  status: "Active",
  repository_url: "https://github.com/example/umya-spreadsheet",
  docs: [
    %{title: "API Documentation", url: "https://docs.example.com/api", description: "Complete API reference"},
    %{title: "User Guide", url: "https://docs.example.com/guide", description: "Step-by-step user guide"},
    %{title: "Examples", url: "https://docs.example.com/examples", description: "Code examples and tutorials"}
  ],
  team: [
    %{name: "Alice Developer", role: "Lead Developer", email: "alice@company.com", slack_id: "U123", team_id: "T456"},
    %{name: "Bob Tester", role: "QA Engineer", email: "bob@company.com", slack_id: "U789", team_id: "T456"},
    %{name: "Carol Manager", role: "Project Manager", email: "carol@company.com", slack_id: nil, team_id: nil}
  ],
  resources: [
    %{name: "CI/CD Pipeline", url: "https://ci.example.com/project", description: "Build and deployment status"},
    %{name: "Issue Tracker", url: "https://issues.example.com/project", description: "Bug reports and feature requests"},
    %{name: "Wiki", url: "https://wiki.example.com/project", description: "Internal project documentation"}
  ]
}

ProjectDocumentationHub.create_hub(project_info)
```

This comprehensive guide covers all aspects of hyperlink management in UmyaSpreadsheet, from basic operations to advanced use cases and real-world examples. The examples demonstrate practical applications and best practices for different scenarios you might encounter when working with Excel hyperlinks.
