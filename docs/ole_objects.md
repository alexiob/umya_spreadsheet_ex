# OLE Objects

This guide explains how to work with OLE (Object Linking and Embedding) objects in Excel spreadsheets using the UmyaSpreadsheet library.

## Overview

OLE objects allow you to embed documents from other applications (like Word documents, PowerPoint presentations, or PDF files) directly into Excel spreadsheets. UmyaSpreadsheet provides comprehensive support for:

- Creating and managing OLE object collections
- Embedding various file types (.docx, .xlsx, .pptx, .pdf, .txt)
- Loading OLE objects from files or binary data
- Managing OLE object properties and metadata
- Retrieving and manipulating embedded objects

## Supported File Types

UmyaSpreadsheet supports the following OLE object types:

- **Word Documents** (.docx) - ProgID: `Word.DocumentMacroEnabled.12`
- **Excel Workbooks** (.xlsx) - ProgID: `Excel.SheetMacroEnabled.12`
- **PowerPoint Presentations** (.pptx) - ProgID: `PowerPoint.ShowMacroEnabled.12`
- **PDF Documents** (.pdf) - ProgID: `AcroExch.Document.DC`
- **Text Files** (.txt) - ProgID: `txtfile`

## Getting Started

### Creating an OLE Objects Collection

First, create a new OLE objects collection for a worksheet:

```elixir
# Create a new spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Create a new OLE objects collection for Sheet1
{:ok, ole_objects} = UmyaSpreadsheet.new_ole_objects()

# Set the OLE objects collection on the worksheet
UmyaSpreadsheet.set_ole_objects(spreadsheet, "Sheet1", ole_objects)
```

### Adding OLE Objects from Files

Add OLE objects by loading them from files:

```elixir
# Add a Word document as an OLE object
case UmyaSpreadsheet.OleObjects.add_ole_object_from_file(ole_objects, "/path/to/document.docx") do
  {:ok, ole_object} ->
    IO.puts("Word document added successfully")
  {:error, reason} ->
    IO.puts("Failed to add document: #{reason}")
end

# Add a PDF document
{:ok, pdf_object} = UmyaSpreadsheet.OleObjects.add_ole_object_from_file(ole_objects, "/path/to/presentation.pdf")

# Add a PowerPoint presentation
{:ok, ppt_object} = UmyaSpreadsheet.OleObjects.add_ole_object_from_file(ole_objects, "/path/to/slides.pptx")
```

### Adding OLE Objects from Binary Data

You can also create OLE objects from binary data:

```elixir
# Read file content as binary
{:ok, binary_data} = File.read("/path/to/document.docx")

# Create OLE object from binary data with explicit ProgID
{:ok, ole_object} = UmyaSpreadsheet.OleObjects.add_ole_object_with_data(
  ole_objects,
  binary_data,
  "Word.DocumentMacroEnabled.12"
)
```

## Working with OLE Objects Collections

### Listing OLE Objects

Get information about all OLE objects in a collection:

```elixir
# Get list of all OLE objects
ole_object_list = UmyaSpreadsheet.OleObjects.get_ole_objects_list(ole_objects)

# Count total objects
count = UmyaSpreadsheet.OleObjects.get_ole_objects_count(ole_objects)
IO.puts("Total OLE objects: #{count}")

# Check if collection has any objects
has_objects = UmyaSpreadsheet.OleObjects.has_ole_objects(ole_objects)
```

### Retrieving Specific OLE Objects

Access individual OLE objects by index:

```elixir
# Get the first OLE object (0-based index)
case UmyaSpreadsheet.OleObjects.get_ole_object_by_index(ole_objects, 0) do
  {:ok, ole_object} ->
    # Work with the OLE object
    prog_id = UmyaSpreadsheet.OleObjects.get_prog_id(ole_object)
    IO.puts("ProgID: #{prog_id}")
  {:error, :not_found} ->
    IO.puts("No OLE object at index 0")
end
```

## Managing OLE Object Properties

### Getting Object Information

Retrieve various properties from OLE objects:

```elixir
# Get the ProgID (identifies the application type)
prog_id = UmyaSpreadsheet.OleObjects.get_prog_id(ole_object)

# Get the file extension
extension = UmyaSpreadsheet.OleObjects.get_extension(ole_object)

# Get any requirements or dependencies
requires = UmyaSpreadsheet.OleObjects.get_requires(ole_object)

# Get the binary data
binary_data = UmyaSpreadsheet.OleObjects.get_ole_object_data(ole_object)

IO.puts("Object Type: #{prog_id}, Extension: #{extension}")
```

### Setting Object Properties

Update OLE object properties:

```elixir
# Set a custom ProgID
UmyaSpreadsheet.OleObjects.set_prog_id(ole_object, "CustomApp.Document.1")

# Set the file extension
UmyaSpreadsheet.OleObjects.set_extension(ole_object, ".custom")

# Set requirements
UmyaSpreadsheet.OleObjects.set_requires(ole_object, "CustomApplication >= 2.0")

# Update the binary data
{:ok, new_data} = File.read("/path/to/updated_file.docx")
UmyaSpreadsheet.OleObjects.set_ole_object_data(ole_object, new_data)
```

## Advanced Operations

### Working with Embedded Object Properties

Access properties specific to embedded objects:

```elixir
# Get embedded object ProgID
embedded_prog_id = UmyaSpreadsheet.OleObjects.get_embedded_object_prog_id(ole_object)

# Get the shape ID associated with the embedded object
shape_id = UmyaSpreadsheet.OleObjects.get_embedded_object_shape_id(ole_object)

IO.puts("Embedded ProgID: #{embedded_prog_id}, Shape ID: #{shape_id}")
```

### File Format Detection

Check the format of OLE objects:

```elixir
# Check if object is in binary format
is_binary = UmyaSpreadsheet.OleObjects.is_binary_format_ole_object(ole_object)

# Check if object is an Excel format
is_excel = UmyaSpreadsheet.OleObjects.is_excel_ole_object(ole_object)

IO.puts("Binary format: #{is_binary}, Excel format: #{is_excel}")
```

### Saving OLE Objects to Files

Extract OLE objects and save them as separate files:

```elixir
# Save an OLE object to a file
case UmyaSpreadsheet.OleObjects.save_ole_object_to_file(ole_object, "/path/to/extracted_document.docx") do
  :ok ->
    IO.puts("OLE object saved successfully")
  {:error, reason} ->
    IO.puts("Failed to save: #{reason}")
end
```

### Helper Functions

Use utility functions for common operations:

```elixir
# Automatically determine ProgID from file extension
prog_id = UmyaSpreadsheet.OleObjects.determine_prog_id_from_extension(".docx")
# Returns: "Word.DocumentMacroEnabled.12"

prog_id = UmyaSpreadsheet.OleObjects.determine_prog_id_from_extension(".pdf")
# Returns: "AcroExch.Document.DC"
```

## Complete Example

Here's a comprehensive example that demonstrates various OLE object operations:

```elixir
defmodule OleExample do
  def create_spreadsheet_with_ole_objects do
    # Create new spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Create OLE objects collection
    {:ok, ole_objects} = UmyaSpreadsheet.new_ole_objects()

    # Add various types of OLE objects
    {:ok, word_object} = UmyaSpreadsheet.OleObjects.add_ole_object_from_file(
      ole_objects,
      "/path/to/contract.docx"
    )

    {:ok, pdf_object} = UmyaSpreadsheet.OleObjects.add_ole_object_from_file(
      ole_objects,
      "/path/to/report.pdf"
    )

    {:ok, ppt_object} = UmyaSpreadsheet.OleObjects.add_ole_object_from_file(
      ole_objects,
      "/path/to/presentation.pptx"
    )

    # Set custom properties
    UmyaSpreadsheet.OleObjects.set_requires(word_object, "Microsoft Word 2016 or later")

    # Set the OLE objects collection on the worksheet
    UmyaSpreadsheet.set_ole_objects(spreadsheet, "Sheet1", ole_objects)

    # Add some context in the spreadsheet
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Embedded Documents")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "Contract (Word)")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "Report (PDF)")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A4", "Presentation (PowerPoint)")

    # Save the spreadsheet
    case UmyaSpreadsheet.write(spreadsheet, "/path/to/spreadsheet_with_ole.xlsx") do
      :ok ->
        IO.puts("Spreadsheet with OLE objects created successfully")
      {:error, reason} ->
        IO.puts("Failed to save: #{reason}")
    end
  end

  def extract_ole_objects_from_spreadsheet do
    # Read existing spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.read("/path/to/spreadsheet_with_ole.xlsx")

    # Get OLE objects from worksheet
    case UmyaSpreadsheet.get_ole_objects(spreadsheet, "Sheet1") do
      {:ok, ole_objects} ->
        count = UmyaSpreadsheet.OleObjects.get_ole_objects_count(ole_objects)
        IO.puts("Found #{count} OLE objects")

        # Extract each object
        for index <- 0..(count - 1) do
          {:ok, ole_object} = UmyaSpreadsheet.OleObjects.get_ole_object_by_index(ole_objects, index)

          prog_id = UmyaSpreadsheet.OleObjects.get_prog_id(ole_object)
          extension = UmyaSpreadsheet.OleObjects.get_extension(ole_object)

          output_file = "/path/to/extracted/object_#{index}#{extension}"

          case UmyaSpreadsheet.OleObjects.save_ole_object_to_file(ole_object, output_file) do
            :ok ->
              IO.puts("Extracted #{prog_id} object to #{output_file}")
            {:error, reason} ->
              IO.puts("Failed to extract object #{index}: #{reason}")
          end
        end

      {:error, :not_found} ->
        IO.puts("No OLE objects found in the worksheet")
    end
  end
end

# Usage
OleExample.create_spreadsheet_with_ole_objects()
OleExample.extract_ole_objects_from_spreadsheet()
```

## Error Handling

OLE object operations can fail for various reasons. Always handle errors appropriately:

```elixir
case UmyaSpreadsheet.OleObjects.add_ole_object_from_file(ole_objects, file_path) do
  {:ok, ole_object} ->
    # Success - work with the object
    IO.puts("OLE object added successfully")

  {:error, :file_not_found} ->
    IO.puts("The specified file does not exist")

  {:error, :unsupported_format} ->
    IO.puts("The file format is not supported for OLE embedding")

  {:error, :invalid_data} ->
    IO.puts("The file contains invalid or corrupted data")

  {:error, reason} ->
    IO.puts("Failed to add OLE object: #{reason}")
end
```

Common error types:

- `:file_not_found` - The specified file path does not exist
- `:unsupported_format` - The file type is not supported for OLE embedding
- `:invalid_data` - The file contains corrupted or invalid data
- `:not_found` - The requested OLE object or index does not exist
- `:error` - General error during processing

## Best Practices

1. **File Size Considerations**: OLE objects embed the entire file content, which can significantly increase spreadsheet size. Consider the impact on file size and loading times.

2. **Application Dependencies**: Recipients of the spreadsheet need the appropriate applications installed to open embedded OLE objects.

3. **Resource Management**: OLE objects collections and individual objects are automatically managed by the library, but ensure you don't hold unnecessary references.

4. **Error Handling**: Always implement proper error handling, especially when working with file operations.

5. **Format Consistency**: Use the helper functions to ensure consistent ProgID assignment based on file extensions.

6. **Testing**: Test OLE object functionality with the actual applications and versions your users will have.

## Limitations

- OLE objects require the target application to be installed on the machine where the Excel file is opened
- Large embedded files can significantly increase spreadsheet file size
- Some file formats may have compatibility issues across different Excel versions
- Complex OLE objects may not display correctly in all Excel environments

For more information about limitations and compatibility, see the [Limitations & Compatibility Guide](limitations.html).
