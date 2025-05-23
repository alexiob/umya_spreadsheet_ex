# Shapes and Drawing

This guide demonstrates how to add shapes, text boxes, and connectors to your Excel spreadsheets using the UmyaSpreadsheet library.

## Overview

UmyaSpreadsheet provides functions to add various drawing elements to your spreadsheets:

- Basic shapes (rectangles, ellipses, triangles, etc.)
- Text boxes for adding annotated text
- Connectors for linking cells or shapes together

These features are useful for:

- Creating diagrams and flowcharts
- Adding annotations and callouts to data
- Building dashboards with visual elements
- Emphasizing important information

## Adding Basic Shapes

UmyaSpreadsheet supports a variety of shapes that can be added to your spreadsheets:

```elixir
# Create a new spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add a blue rectangle at cell A1
UmyaSpreadsheet.add_shape(
  spreadsheet,
  "Sheet1",
  "A1",                # Cell position
  "rectangle",         # Shape type
  200,                 # Width in pixels
  100,                 # Height in pixels
  "blue",              # Fill color
  "black",             # Outline color
  1.0                  # Outline width in points
)

# Add a red circle at cell B5
UmyaSpreadsheet.add_shape(
  spreadsheet,
  "Sheet1",
  "B5",
  "ellipse",           # Use "ellipse" for circles and ovals
  150,                 # Equal width and height creates a circle
  150,
  "red",
  "black",
  1.0
)

# Save the spreadsheet
UmyaSpreadsheet.write(spreadsheet, "spreadsheet_with_shapes.xlsx")
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

## Adding Text Boxes

Text boxes are special shapes that can contain text, useful for adding annotations or callouts:

```elixir
# Add a text box with white background and black text at cell D10
UmyaSpreadsheet.add_text_box(
  spreadsheet,
  "Sheet1",
  "D10",               # Cell position
  "Important Note",    # Text content
  200,                 # Width in pixels
  100,                 # Height in pixels
  "white",             # Background color
  "black",             # Text color
  "#888888",           # Border color (using hex code)
  1.0                  # Border width in points
)
```

## Adding Connectors

Connectors are lines that connect two cells, useful for creating flowcharts and diagrams:

```elixir
# Add a connector line from cell A1 to cell C3
UmyaSpreadsheet.add_connector(
  spreadsheet,
  "Sheet1",
  "A1",                # Starting cell
  "C3",                # Ending cell
  "red",               # Line color
  1.5                  # Line width in points
)
```

## Color Specification

All shape functions accept colors specified in two formats:

1. **Named colors**: Common color names like "red", "blue", "green", "white", "black", etc.
2. **Hex codes**: Standard hex color codes, with or without the # prefix (e.g., "#FF0000" or "FF0000" for red)

## Creating Diagrams

You can combine shapes, text boxes, and connectors to create diagrams and flowcharts:

```elixir
# Create a simple flowchart
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add shapes for the flowchart nodes
UmyaSpreadsheet.add_shape(spreadsheet, "Sheet1", "B2", "rectangle", 150, 80, "#E6F2FF", "black", 1.0)
UmyaSpreadsheet.add_shape(spreadsheet, "Sheet1", "B6", "diamond", 150, 100, "#FFE6E6", "black", 1.0)
UmyaSpreadsheet.add_shape(spreadsheet, "Sheet1", "B10", "rectangle", 150, 80, "#E6FFE6", "black", 1.0)

# Add text boxes for the node labels
UmyaSpreadsheet.add_text_box(spreadsheet, "Sheet1", "B2", "Start", 150, 80, "transparent", "black", "transparent", 0.0)
UmyaSpreadsheet.add_text_box(spreadsheet, "Sheet1", "B6", "Decision", 150, 100, "transparent", "black", "transparent", 0.0)
UmyaSpreadsheet.add_text_box(spreadsheet, "Sheet1", "B10", "End", 150, 80, "transparent", "black", "transparent", 0.0)

# Add connectors between the nodes
UmyaSpreadsheet.add_connector(spreadsheet, "Sheet1", "B4", "B6", "black", 1.0)
UmyaSpreadsheet.add_connector(spreadsheet, "Sheet1", "B8", "B10", "black", 1.0)

# Save the flowchart
UmyaSpreadsheet.write(spreadsheet, "flowchart.xlsx")
```

## Tips for Working with Shapes

1. **Positioning**: Cell coordinates determine where shapes are positioned. The shape will be anchored to the top-left corner of the specified cell.

2. **Sizing**: Width and height are specified in pixels. Excel uses a resolution of approximately 96 pixels per inch.

3. **Layering**: Shapes are added in the order they are created. Later shapes appear on top of earlier shapes.

4. **Transparency**: You can use "transparent" as a color value to create shapes without fill or outline.

5. **Text Boxes**: For multi-line text in text boxes, use the newline character ("\n") to create line breaks.

## See Also

- [Image Handling](image_handling.html) - For adding images to spreadsheets
- [Styling and Formatting](styling_and_formatting.html) - For cell and sheet styling options
- [Sheet Operations](sheet_operations.html) - For managing worksheets
- [Thread Safety](thread_safety.html) - For information about concurrent operations
