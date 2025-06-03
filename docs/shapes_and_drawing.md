# Drawing and Shape Management

This guide explains how to add and retrieve shapes, text boxes, and connectors in Excel spreadsheets using UmyaSpreadsheet.

## Introduction to Drawings

Excel supports various drawing objects such as shapes, text boxes, and connector lines that help create visual elements and diagrams in spreadsheets. UmyaSpreadsheet provides functions to both add and retrieve these objects.

### Supported Drawing Types

UmyaSpreadsheet supports these drawing types:

- **Shapes**: Various geometric shapes including rectangles, ellipses, triangles, diamonds, etc.
- **Text Boxes**: Text containers with customizable appearance
- **Connectors**: Lines connecting cells or shapes for creating flowcharts and diagrams

These features are useful for:

- Creating diagrams and flowcharts
- Adding annotations and callouts to data
- Building dashboards with visual elements
- Emphasizing important information

## Adding Drawing Objects

### Adding Shapes

```elixir
alias UmyaSpreadsheet.Drawing

# Add a blue rectangle at cell A1
Drawing.add_shape(
  spreadsheet,
  "Sheet1",
  "A1",                # Cell address where shape will be placed
  "rectangle",         # Shape type
  200.0,               # Width in pixels
  100.0,               # Height in pixels
  "blue",              # Fill color
  "black",             # Outline color
  1.0                  # Outline width in points
)
```

### Supported Shape Types

UmyaSpreadsheet supports the following shape types:

| Shape Type         | Description                           |
|--------------------|---------------------------------------|
| `"rectangle"`      | Rectangle shape                       |
| `"ellipse"`        | Ellipse shape (also "oval", "circle") |
| `"rounded_rectangle"` | Rectangle with rounded corners     |
| `"triangle"`       | Triangle shape                        |
| `"right_triangle"` | Right triangle shape                  |
| `"pentagon"`       | Pentagon shape                        |
| `"hexagon"`        | Hexagon shape                         |
| `"octagon"`        | Octagon shape                         |
| `"trapezoid"`      | Trapezoid shape                       |
| `"diamond"`        | Diamond shape                         |
| `"arrow"`          | Arrow shape                           |
| `"line"`           | Line shape                            |

### Adding Text Boxes

```elixir
# Add a white text box with black text at cell B2
Drawing.add_text_box(
  spreadsheet,
  "Sheet1",
  "B2",                # Cell address where text box will be placed
  "Hello World",       # Text content
  150.0,               # Width in pixels
  75.0,                # Height in pixels
  "white",             # Background color
  "black",             # Text color
  "gray",              # Border color
  1.0                  # Border width in points
)
```

### Adding Connectors

```elixir
# Add a red connector line from cell C3 to cell D5
Drawing.add_connector(
  spreadsheet,
  "Sheet1",
  "C3",                # Starting cell
  "D5",                # Ending cell
  "red",               # Line color
  1.5                  # Line width in points
)
```

## Retrieving Drawing Objects

UmyaSpreadsheet provides getter functions to retrieve and analyze drawing objects in a spreadsheet.

### Getting All Shapes

```elixir
# Get all shapes in a sheet
{:ok, shapes} = Drawing.get_shapes(spreadsheet, "Sheet1")

# Get shapes in a specific range
{:ok, shapes_in_range} = Drawing.get_shapes(spreadsheet, "Sheet1", "A1:E10")

# Each shape in the result contains these properties:
# - type: The shape type ("rectangle", "ellipse", etc.)
# - cell: The cell address where the shape is placed
# - width: The width of the shape in pixels
# - height: The height of the shape in pixels
# - fill_color: The fill color of the shape as a hex code
# - outline_color: The outline color as a hex code
# - outline_width: The width of the outline in points
```

### Getting Text Boxes

```elixir
# Get all text boxes in a sheet
{:ok, text_boxes} = Drawing.get_text_boxes(spreadsheet, "Sheet1")

# Get text boxes in a specific range
{:ok, text_boxes_in_range} = Drawing.get_text_boxes(spreadsheet, "Sheet1", "B2:C5")

# Each text box in the result contains these properties:
# - cell: The cell address where the text box is placed
# - text: The text content of the text box
# - width: The width of the text box in pixels
# - height: The height of the text box in pixels
# - fill_color: The background color as a hex code
# - text_color: The text color as a hex code
# - outline_color: The border color as a hex code
# - outline_width: The width of the border in points
```

### Getting Connectors

```elixir
# Get all connectors in a sheet
{:ok, connectors} = Drawing.get_connectors(spreadsheet, "Sheet1")

# Get connectors in a specific range (matches if either end is in range)
{:ok, connectors_in_range} = Drawing.get_connectors(spreadsheet, "Sheet1", "C3:D5")

# Each connector in the result contains these properties:
# - from_cell: The starting cell address
# - to_cell: The ending cell address
# - line_color: The connector color as a hex code
# - line_width: The width of the line in points
```

### Utility Functions

```elixir
# Check if a sheet has any drawing objects
{:ok, has_drawings} = Drawing.has_drawing_objects(spreadsheet, "Sheet1")

# Count drawing objects in a sheet
{:ok, count} = Drawing.count_drawing_objects(spreadsheet, "Sheet1")

# Count drawing objects in a specific range
{:ok, count_in_range} = Drawing.count_drawing_objects(spreadsheet, "Sheet1", "A1:E10")
```

## Complete Example

Here's a complete example showing how to create and analyze a diagram with shapes and connectors:

```elixir
alias UmyaSpreadsheet.Drawing

# Create a new spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add title
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Process Diagram")
UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1", true)
UmyaSpreadsheet.set_font_size(spreadsheet, "Sheet1", "A1", 14)

# Add shapes for a flowchart
Drawing.add_text_box(spreadsheet, "Sheet1", "B3", "Start", 100, 40, "#D9EAD3", "black", "#274E13", 1.0)
Drawing.add_text_box(spreadsheet, "Sheet1", "B5", "Process 1", 100, 40, "#FCE5CD", "black", "#B45F06", 1.0)
Drawing.add_shape(spreadsheet, "Sheet1", "B7", "diamond", 100, 80, "#CFE2F3", "black", 1.0)
Drawing.add_text_box(spreadsheet, "Sheet1", "E7", "Process 2", 100, 40, "#FCE5CD", "black", "#B45F06", 1.0)
Drawing.add_text_box(spreadsheet, "Sheet1", "B9", "End", 100, 40, "#F4CCCC", "black", "#990000", 1.0)

# Add connectors to create a flowchart
Drawing.add_connector(spreadsheet, "Sheet1", "B3", "B5", "black", 1.0)
Drawing.add_connector(spreadsheet, "Sheet1", "B5", "B7", "black", 1.0)
Drawing.add_connector(spreadsheet, "Sheet1", "B7", "E7", "black", 1.0)
Drawing.add_connector(spreadsheet, "Sheet1", "B7", "B9", "black", 1.0)
Drawing.add_connector(spreadsheet, "Sheet1", "E7", "B9", "black", 1.0)

# Save the spreadsheet
UmyaSpreadsheet.write_file(spreadsheet, "flowchart.xlsx")

# Analyze the drawing objects
{:ok, shape_count} = Drawing.count_drawing_objects(spreadsheet, "Sheet1")
IO.puts("Total number of drawing objects: #{shape_count}")

{:ok, connectors} = Drawing.get_connectors(spreadsheet, "Sheet1")
IO.puts("Number of connector lines: #{length(connectors)}")

{:ok, text_boxes} = Drawing.get_text_boxes(spreadsheet, "Sheet1")
IO.puts("Number of text boxes: #{length(text_boxes)}")
```

## Color Specification

All shape functions accept colors specified in two formats:

1. **Named colors**: Common color names like "red", "blue", "green", "white", "black", etc.
2. **Hex codes**: Standard hex color codes, with or without the # prefix (e.g., "#FF0000" or "FF0000" for red)

## Tips for Working with Drawing Objects

1. **Cell Positioning**: Drawing objects are anchored to cells but can extend beyond them

2. **Color Values**: Colors can be specified as named colors ("red", "blue") or hex codes ("#FF0000")

3. **Sizing**: Width and height are specified in pixels. Excel uses a resolution of approximately 96 pixels per inch.

4. **Layering**: Shapes are added in the order they are created. Later shapes appear on top of earlier shapes.

5. **Transparency**: You can use "transparent" as a color value to create shapes without fill or outline.

6. **Text Boxes**: For multi-line text in text boxes, use the newline character ("\n") to create line breaks.

7. **Filtering**: Use the optional cell range parameter in getter functions to focus on specific areas of a sheet

8. **Optimization**: For large sheets with many drawings, filter by range to improve performance

## See Also

- [Image Handling](image_handling.html) - For adding images to spreadsheets
- [Styling and Formatting](styling_and_formatting.html) - For cell and sheet styling options
- [Sheet Operations](sheet_operations.html) - For managing worksheets
- [Thread Safety](thread_safety.html) - For information about concurrent operations
