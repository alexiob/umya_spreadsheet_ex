# Image Handling

This guide explains how to work with images in Excel spreadsheets using the UmyaSpreadsheet library.

## Overview

UmyaSpreadsheet provides comprehensive image handling functions to:

- **Add images** to specific cells with positioning control
- **Retrieve and download** images from existing spreadsheets
- **Replace existing images** with new ones
- **Get image dimensions** (width and height)
- **List all images** in a worksheet with their positions
- **Get detailed image information** including names, positions, and dimensions
- **Manage image positioning and sizing** within cells

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

## Getting Image Information

### Get Image Dimensions

You can retrieve the dimensions of an image at a specific cell:

```elixir
# Get the width and height of an image
case UmyaSpreadsheet.get_image_dimensions(spreadsheet, "Sheet1", "A1") do
  {:ok, {width, height}} ->
    IO.puts("Image dimensions: #{width}x#{height} pixels")
  {:error, reason} ->
    IO.puts("Failed to get image dimensions: #{reason}")
end
```

### Get Comprehensive Image Information

To get detailed information about an image including its name, position, and dimensions:

```elixir
# Get comprehensive image information
case UmyaSpreadsheet.get_image_info(spreadsheet, "Sheet1", "A1") do
  {:ok, {name, position, width, height}} ->
    IO.puts("Image '#{name}' at position #{position}, dimensions: #{width}x#{height} pixels")
  {:error, reason} ->
    IO.puts("Failed to get image info: #{reason}")
end
```

### List All Images in a Sheet

To get a list of all images in a specific worksheet:

```elixir
# List all images in the sheet
case UmyaSpreadsheet.list_images(spreadsheet, "Sheet1") do
  {:ok, images} ->
    IO.puts("Found #{length(images)} images in Sheet1:")
    Enum.each(images, fn {coordinate, image_name} ->
      IO.puts("  - '#{image_name}' at #{coordinate}")
    end)
  {:error, reason} ->
    IO.puts("Failed to list images: #{reason}")
end
```

## Working with Multiple Images

Here's a comprehensive example showing how to work with multiple images:

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add multiple images
UmyaSpreadsheet.add_image(spreadsheet, "Sheet1", "A1", "/path/to/logo.png")
UmyaSpreadsheet.add_image(spreadsheet, "Sheet1", "C1", "/path/to/chart.png")
UmyaSpreadsheet.add_image(spreadsheet, "Sheet1", "E1", "/path/to/diagram.png")

# List all images to verify they were added
{:ok, images} = UmyaSpreadsheet.list_images(spreadsheet, "Sheet1")
IO.puts("Added #{length(images)} images to the spreadsheet")

# Get detailed information for each image
Enum.each(images, fn {coordinate, _name} ->
  {:ok, {name, pos, width, height}} = UmyaSpreadsheet.get_image_info(spreadsheet, "Sheet1", coordinate)
  IO.puts("Image '#{name}' at #{pos}: #{width}x#{height} pixels")
end)

# Save the spreadsheet
UmyaSpreadsheet.write(spreadsheet, "multiple_images.xlsx")
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

By default, images are positioned at the top-left corner of the specified cell. The library provides the basic `add_image/4` function for simple image insertion:

```elixir
# Add an image to a specific cell
UmyaSpreadsheet.add_image(spreadsheet, "Sheet1", "A1", "/path/to/image.png")
```

**Note**: Advanced positioning and sizing options (such as custom width, height, and offset parameters) may be available in future versions of the library. Currently, images are inserted with their original dimensions at the cell's top-left corner.

For precise control over image appearance in your Excel files:

1. **Pre-size your images** to the desired dimensions before adding them to the spreadsheet
2. **Use appropriate image formats** (PNG, JPEG, GIF) that are well-supported by Excel
3. **Consider cell dimensions** when choosing image sizes for optimal appearance
4. **Test the output** in Excel to ensure images appear as expected

## Best Practices

When working with images in spreadsheets:

1. Use appropriate image formats (PNG, JPEG, GIF) that are well-supported by Excel
2. Be mindful of image file sizes to keep the spreadsheet file size manageable
3. Set explicit image dimensions when possible for consistent appearance
4. Consider using relative paths when working with images in a project structure

## See Also

- [Styling and Formatting](styling_and_formatting.html) - For overall spreadsheet styling
- [Sheet Operations](sheet_operations.html) - For managing sheets that contain images
