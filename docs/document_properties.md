# Document Properties Guide

This guide covers how to work with document properties (metadata) in UmyaSpreadsheet. Document properties allow you to store additional information about your spreadsheets, such as author information, project details, and custom business data.

## Overview

Excel files can store various types of metadata through document properties:

1. **Core Properties** - Standard metadata fields like title, author, description
2. **Custom Properties** - User-defined key-value pairs for application-specific data

UmyaSpreadsheet provides a complete API to read, write, and manage both types of properties.

## Core Document Properties

Core properties are standard metadata fields supported by Excel. They provide basic information about the document.

### Setting Core Properties

```elixir
alias UmyaSpreadsheet.DocumentProperties

# Open an existing spreadsheet or create a new one
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Set individual core properties
DocumentProperties.set_title(spreadsheet, "Q2 Financial Report")
DocumentProperties.set_description(spreadsheet, "Quarterly financial analysis")
DocumentProperties.set_subject(spreadsheet, "Finance")
DocumentProperties.set_creator(spreadsheet, "Finance Team")
DocumentProperties.set_company(spreadsheet, "Acme Corporation")
DocumentProperties.set_category(spreadsheet, "Financial Reports")
DocumentProperties.set_manager(spreadsheet, "Jane Smith")
DocumentProperties.set_last_modified_by(spreadsheet, "John Doe")
DocumentProperties.set_created(spreadsheet, "2025-05-28T09:00:00Z")
DocumentProperties.set_modified(spreadsheet, "2025-05-28T14:30:00Z")
```

### Reading Core Properties

```elixir
# Get individual core properties
{:ok, title} = DocumentProperties.get_title(spreadsheet)
{:ok, creator} = DocumentProperties.get_creator(spreadsheet)
{:ok, description} = DocumentProperties.get_description(spreadsheet)

# Get all core properties at once
{:ok, properties} = DocumentProperties.get_all_properties(spreadsheet)
IO.puts("Title: #{properties.title}")
IO.puts("Creator: #{properties.creator}")
IO.puts("Created: #{properties.created}")
```

### Setting Multiple Properties at Once

```elixir
# Set multiple core properties in a single operation
properties = %{
  title: "Q2 Financial Report",
  description: "Quarterly financial analysis",
  subject: "Finance",
  creator: "Finance Team",
  company: "Acme Corporation"
}

DocumentProperties.set_properties(spreadsheet, properties)
```

## Custom Document Properties

Custom properties allow you to store arbitrary application-specific data as key-value pairs.

### Setting Custom Properties

UmyaSpreadsheet supports various data types for custom properties:

```elixir
# String properties
DocumentProperties.set_custom_property_string(spreadsheet, "ProjectCode", "FIN-2025-Q2")

# Numeric properties
DocumentProperties.set_custom_property_number(spreadsheet, "Version", 1.5)
DocumentProperties.set_custom_property_number(spreadsheet, "DocumentID", 12345)

# Boolean properties
DocumentProperties.set_custom_property_bool(spreadsheet, "IsConfidential", true)

# Date properties
DocumentProperties.set_custom_property_date(spreadsheet, "ReviewDate", "2025-06-15T00:00:00Z")
```

### Automatic Type Detection

For convenience, you can use the generic `set_custom_property` function which automatically determines the data type:

```elixir
# Automatically detects types
DocumentProperties.set_custom_property(spreadsheet, "ProjectName", "Financial Analysis")  # String
DocumentProperties.set_custom_property(spreadsheet, "Version", 2.0)                      # Number
DocumentProperties.set_custom_property(spreadsheet, "IsApproved", true)                  # Boolean
```

### Reading Custom Properties

```elixir
# Read a custom property (returns the value with proper type)
{:ok, project_code} = DocumentProperties.get_custom_property(spreadsheet, "ProjectCode")
{:ok, version} = DocumentProperties.get_custom_property(spreadsheet, "Version")
{:ok, is_confidential} = DocumentProperties.get_custom_property(spreadsheet, "IsConfidential")

# Check if a property exists
{:ok, exists} = DocumentProperties.has_custom_property(spreadsheet, "ProjectCode")

# Get property names and count
{:ok, names} = DocumentProperties.get_custom_property_names(spreadsheet)
{:ok, count} = DocumentProperties.get_custom_properties_count(spreadsheet)
```

### Managing Custom Properties

```elixir
# Remove a single custom property
DocumentProperties.remove_custom_property(spreadsheet, "TemporaryData")

# Clear all custom properties
DocumentProperties.clear_custom_properties(spreadsheet)
```

## Persistence

Document properties are automatically saved when writing the spreadsheet file and loaded when reading a file:

```elixir
# Save the spreadsheet with properties
UmyaSpreadsheet.write(spreadsheet, "financial_report.xlsx")

# Read the spreadsheet with its properties
{:ok, loaded_spreadsheet} = UmyaSpreadsheet.read("financial_report.xlsx")

# Access properties from the loaded spreadsheet
{:ok, title} = DocumentProperties.get_title(loaded_spreadsheet)
```

## Edge Cases and Best Practices

### Date Handling

When working with dates in document properties:

```elixir
# ISO 8601 format is recommended
DocumentProperties.set_created(spreadsheet, "2025-05-28T09:00:00Z")

# Simple date format is also supported
DocumentProperties.set_custom_property_date(spreadsheet, "DueDate", "2025-06-30")
```

### Unicode Support

Document properties fully support Unicode characters:

```elixir
# Unicode characters are fully supported
DocumentProperties.set_title(spreadsheet, "財務報告 2025")
DocumentProperties.set_custom_property_string(spreadsheet, "Location", "São Paulo 🇧🇷")
```

### Large Numbers

When working with very large numbers:

```elixir
# Very large numbers may be stored as strings in Excel
large_number = 1.7976931348623157e+308
DocumentProperties.set_custom_property_number(spreadsheet, "LargeValue", large_number)

# When reading back, compare with small tolerance for floating point precision
{:ok, value} = DocumentProperties.get_custom_property(spreadsheet, "LargeValue")
```

## Complete API Reference

For the complete API reference, please check the documentation for `UmyaSpreadsheet.DocumentProperties` module.

## Common Issues and Troubleshooting

### Property Not Found

If you receive an `{:error, :not_found}` when trying to read a property, make sure:

1. You're using the exact property name (case-sensitive)
2. The property exists in the document
3. You're using the correct property type function

### Error Handling

Always handle potential errors when working with document properties:

```elixir
case DocumentProperties.get_custom_property(spreadsheet, "Status") do
  {:ok, value} ->
    # Use the value
    IO.puts("Status: #{value}")
  {:error, :not_found} ->
    # Property doesn't exist
    IO.puts("Status property not found")
  {:error, reason} ->
    # Other error
    IO.puts("Error retrieving status: #{inspect(reason)}")
end
```

### Performance Considerations

Document properties are lightweight metadata and typically have minimal performance impact, even with hundreds of properties. However, for very large documents with many properties, consider:

1. Using bulk operations where available
2. Limiting property string lengths when possible
3. Using numeric or boolean properties instead of strings for frequently updated values

## Conclusion

Document properties provide a powerful way to add metadata to your Excel files. Whether you're storing standard information like titles and authors or custom application-specific data, UmyaSpreadsheet's DocumentProperties module gives you complete control over this important aspect of spreadsheet management.
