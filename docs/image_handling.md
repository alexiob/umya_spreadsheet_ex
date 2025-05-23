# Image Handling

This guide explains how to work with images in Excel spreadsheets using the UmyaSpreadsheet library.

## Overview

UmyaSpreadsheet provides functions to:

- Add images to specific cells
- Retrieve images from existing spreadsheets
- Replace existing images
- Manage image positioning and sizing

## Adding Images

To add an image to a spreadsheet:

```elixir
# Create a new spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add an image to cell A1 of Sheet1
UmyaSpreadsheet.add_image(spreadsheet, "Sheet1", "A1", "/path/to/image.png")

# Save the spreadsheet
UmyaSpreadsheet.write(spreadsheet, "spreadsheet_with_image.xlsx")
```

## Retrieving Images

To retrieve an image from an existing spreadsheet:

```elixir
# Read a spreadsheet containing images
{:ok, spreadsheet} = UmyaSpreadsheet.read("spreadsheet_with_image.xlsx")

# Download the image from cell A1 to a file
UmyaSpreadsheet.download_image(spreadsheet, "Sheet1", "A1", "output_image.png")
```

## Changing Images

To replace an existing image in a spreadsheet:

```elixir
# Read a spreadsheet containing images
{:ok, spreadsheet} = UmyaSpreadsheet.read("spreadsheet_with_image.xlsx")

# Change the image at cell A1
UmyaSpreadsheet.change_image(spreadsheet, "Sheet1", "A1", "/path/to/new_image.png")

# Save the updated spreadsheet
UmyaSpreadsheet.write(spreadsheet, "updated_spreadsheet.xlsx")
```

## Error Handling

All image handling functions return `:ok` on success and `{:error, reason}` on failure:

```elixir
case UmyaSpreadsheet.add_image(spreadsheet, "Sheet1", "A1", "/path/to/image.png") do
  :ok ->
    IO.puts("Image added successfully")
  {:error, :not_found} ->
    IO.puts("Sheet not found")
  {:error, _} ->
    IO.puts("Failed to add image")
end
```

Common error reasons:

- `:not_found` - The specified sheet or cell doesn't exist
- `:invalid_format` - The image file format is not supported
- `:file_not_found` - The image file does not exist or cannot be accessed
- `:error` - General error during processing

## Image Positioning and Sizing

By default, images are positioned at the top-left corner of the specified cell. You can control positioning and sizing with additional parameters:

```elixir
# Add an image with custom width and height
UmyaSpreadsheet.add_image(
  spreadsheet,
  "Sheet1",
  "A1",
  "/path/to/image.png",
  %{width: 300, height: 200}
)

# Add an image with positioning offset
UmyaSpreadsheet.add_image(
  spreadsheet,
  "Sheet1",
  "A1",
  "/path/to/image.png",
  %{offset_x: 10, offset_y: 5}
)
```

## Best Practices

When working with images in spreadsheets:

1. Use appropriate image formats (PNG, JPEG, GIF) that are well-supported by Excel
2. Be mindful of image file sizes to keep the spreadsheet file size manageable
3. Set explicit image dimensions when possible for consistent appearance
4. Consider using relative paths when working with images in a project structure

## See Also

- [Styling and Formatting](styling_and_formatting.html) - For overall spreadsheet styling
- [Sheet Operations](sheet_operations.html) - For managing sheets that contain images
