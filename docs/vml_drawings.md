# VML Drawings Guide

Vector Markup Language (VML) is a legacy format used in Excel for creating shapes, comments, and embedded drawing objects. While modern Excel primarily uses DrawingML for new drawings, VML is still widely used for backward compatibility and certain types of shapes. UmyaSpreadsheet provides comprehensive support for creating and manipulating VML shapes in Excel files.

## Table of Contents

1. [Overview](#overview)
2. [Quick Start](#quick-start)
3. [API Reference](#api-reference)
4. [Shape Types](#shape-types)
5. [Style Properties](#style-properties)
6. [Fill and Stroke Properties](#fill-and-stroke-properties)
7. [Examples](#examples)
8. [Best Practices](#best-practices)
9. [Limitations](#limitations)
10. [Troubleshooting](#troubleshooting)

## Overview

VML (Vector Markup Language) was Microsoft's early standard for vector graphics in web and Office applications. In Excel, VML is used for:

- **Legacy Shapes**: Basic geometric shapes like rectangles, ovals, and lines
- **Comment Callouts**: Visual indicators for cell comments
- **Form Controls**: Legacy form controls and buttons
- **Custom Drawing Objects**: User-defined shapes and graphics

### Key Features

- **Shape Creation**: Create basic geometric shapes in worksheets
- **Style Control**: Position, size, and appearance customization
- **Fill Properties**: Control shape fill visibility and colors
- **Stroke Properties**: Outline appearance, colors, and thickness
- **Legacy Compatibility**: Support for older Excel file formats

### When to Use VML

- **Legacy Support**: When working with older Excel files that use VML
- **Simple Shapes**: For basic geometric shapes and annotations
- **Comment Integration**: When creating custom comment indicators
- **Backward Compatibility**: Ensuring compatibility with older Excel versions

## Quick Start

Here's a simple example of creating a VML shape:

```elixir
# Create a new spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Create a basic rectangle shape
:ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "Sheet1", "rectangle1")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, "Sheet1", "rectangle1", "rect")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_style(spreadsheet, "Sheet1", "rectangle1",
  "position:absolute;left:100pt;top:100pt;width:200pt;height:100pt")

# Set fill properties
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_filled(spreadsheet, "Sheet1", "rectangle1", true)
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(spreadsheet, "Sheet1", "rectangle1", "#CCFFCC")

# Set stroke properties
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroked(spreadsheet, "Sheet1", "rectangle1", true)
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_color(spreadsheet, "Sheet1", "rectangle1", "#009900")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_weight(spreadsheet, "Sheet1", "rectangle1", "2pt")

# Save the file
:ok = UmyaSpreadsheet.write(spreadsheet, "shapes_example.xlsx")
```

## API Reference

### Core Functions

#### `create_shape/3`

Creates a new VML shape in the specified worksheet.

```elixir
@spec create_shape(Spreadsheet.t(), String.t(), String.t()) :: :ok | {:error, String.t()}
```

**Parameters:**

- `spreadsheet` - A spreadsheet struct
- `sheet_name` - The name of the worksheet
- `shape_id` - Unique identifier for the shape

**Returns:**

- `:ok` on success
- `{:error, reason}` on failure

**Example:**

```elixir
:ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "Sheet1", "myShape")
```

#### `set_shape_type/4`

Sets the type of a VML shape.

```elixir
@spec set_shape_type(Spreadsheet.t(), String.t(), String.t(), String.t()) :: :ok | {:error, String.t()}
```

**Parameters:**

- `spreadsheet` - A spreadsheet struct
- `sheet_name` - The name of the worksheet
- `shape_id` - Unique identifier for the shape
- `shape_type` - The shape type (see [Shape Types](#shape-types))

**Example:**

```elixir
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, "Sheet1", "myShape", "oval")
```

#### `set_shape_style/4`

Sets the CSS style for positioning and sizing a VML shape.

```elixir
@spec set_shape_style(Spreadsheet.t(), String.t(), String.t(), String.t()) :: :ok | {:error, String.t()}
```

**Parameters:**

- `spreadsheet` - A spreadsheet struct
- `sheet_name` - The name of the worksheet
- `shape_id` - Unique identifier for the shape
- `style` - CSS style string for positioning and dimensions

**Example:**

```elixir
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_style(spreadsheet, "Sheet1", "myShape",
  "position:absolute;left:50pt;top:75pt;width:150pt;height:100pt")
```

### Fill Properties

#### `set_shape_filled/4`

Controls whether a VML shape is filled.

```elixir
@spec set_shape_filled(Spreadsheet.t(), String.t(), String.t(), boolean()) :: :ok | {:error, String.t()}
```

**Parameters:**

- `filled` - `true` to enable fill, `false` to disable

**Example:**

```elixir
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_filled(spreadsheet, "Sheet1", "myShape", true)
```

#### `set_shape_fill_color/4`

Sets the fill color of a VML shape.

```elixir
@spec set_shape_fill_color(Spreadsheet.t(), String.t(), String.t(), String.t()) :: :ok | {:error, String.t()}
```

**Parameters:**

- `fill_color` - The fill color in hex format (e.g., "#FF0000")

**Example:**

```elixir
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(spreadsheet, "Sheet1", "myShape", "#FFCC99")
```

### Stroke Properties

#### `set_shape_stroked/4`

Controls whether a VML shape has a stroke (outline).

```elixir
@spec set_shape_stroked(Spreadsheet.t(), String.t(), String.t(), boolean()) :: :ok | {:error, String.t()}
```

**Parameters:**

- `stroked` - `true` to enable stroke, `false` to disable

**Example:**

```elixir
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroked(spreadsheet, "Sheet1", "myShape", true)
```

#### `set_shape_stroke_color/4`

Sets the stroke (outline) color of a VML shape.

```elixir
@spec set_shape_stroke_color(Spreadsheet.t(), String.t(), String.t(), String.t()) :: :ok | {:error, String.t()}
```

**Parameters:**

- `stroke_color` - The stroke color in hex format

**Example:**

```elixir
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_color(spreadsheet, "Sheet1", "myShape", "#0066CC")
```

#### `set_shape_stroke_weight/4`

Sets the stroke (outline) thickness of a VML shape.

```elixir
@spec set_shape_stroke_weight(Spreadsheet.t(), String.t(), String.t(), String.t()) :: :ok | {:error, String.t()}
```

**Parameters:**

- `stroke_weight` - The stroke thickness (e.g., "1pt", "2px", "0.5em")

**Example:**

```elixir
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_weight(spreadsheet, "Sheet1", "myShape", "3pt")
```

## Shape Types

VML supports various shape types for different geometric forms:

### Basic Shapes

| Shape Type | Description | Use Cases |
|------------|-------------|-----------|
| `"rect"` | Rectangle | Boxes, containers, highlights |
| `"oval"` | Oval/Circle | Buttons, indicators, callouts |
| `"line"` | Straight line | Connectors, dividers, arrows |
| `"roundrect"` | Rounded rectangle | Modern UI elements, buttons |

### Advanced Shapes

| Shape Type | Description | Use Cases |
|------------|-------------|-----------|
| `"polyline"` | Multi-segment line | Complex paths, charts |
| `"arc"` | Curved arc | Decorative elements, connectors |
| `"curve"` | Bezier curve | Smooth curved lines |
| `"polygon"` | Multi-sided polygon | Custom geometric shapes |

### Shape Type Examples

```elixir
# Rectangle
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, "Sheet1", "box", "rect")

# Circle/Oval
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, "Sheet1", "circle", "oval")

# Line
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, "Sheet1", "line", "line")

# Rounded Rectangle
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, "Sheet1", "button", "roundrect")
```

## Style Properties

VML shapes use CSS-like style strings for positioning and sizing. The style string contains semicolon-separated property-value pairs.

### Position Properties

| Property | Description | Units | Example |
|----------|-------------|-------|---------|
| `position` | Positioning mode | `absolute`, `relative` | `position:absolute` |
| `left` | Horizontal position | `pt`, `px`, `em`, `%` | `left:100pt` |
| `top` | Vertical position | `pt`, `px`, `em`, `%` | `top:150pt` |

### Size Properties

| Property | Description | Units | Example |
|----------|-------------|-------|---------|
| `width` | Shape width | `pt`, `px`, `em`, `%` | `width:200pt` |
| `height` | Shape height | `pt`, `px`, `em`, `%` | `height:100pt` |

### Style Examples

```elixir
# Basic positioning and sizing
style = "position:absolute;left:50pt;top:75pt;width:150pt;height:100pt"
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_style(spreadsheet, "Sheet1", "shape1", style)

# Using different units
style = "position:absolute;left:2em;top:1em;width:300px;height:150px"
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_style(spreadsheet, "Sheet1", "shape2", style)

# Percentage-based sizing (relative to container)
style = "position:absolute;left:10%;top:5%;width:80%;height:20%"
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_style(spreadsheet, "Sheet1", "shape3", style)
```

### Unit Types

- **pt (points)**: 1/72 of an inch, commonly used in print
- **px (pixels)**: Screen pixels, device-dependent
- **em**: Relative to font size
- **% (percentage)**: Relative to container dimensions

## Fill and Stroke Properties

### Fill Properties

Fill properties control the interior appearance of shapes:

```elixir
# Enable fill and set color
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_filled(spreadsheet, "Sheet1", "shape1", true)
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(spreadsheet, "Sheet1", "shape1", "#FFEECC")

# Disable fill for transparent shape
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_filled(spreadsheet, "Sheet1", "shape2", false)
```

### Stroke Properties

Stroke properties control the outline appearance:

```elixir
# Enable stroke with color and thickness
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroked(spreadsheet, "Sheet1", "shape1", true)
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_color(spreadsheet, "Sheet1", "shape1", "#336699")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_weight(spreadsheet, "Sheet1", "shape1", "2pt")

# Disable stroke for borderless shape
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroked(spreadsheet, "Sheet1", "shape2", false)
```

### Color Formats

VML supports various color formats:

- **Hex Colors**: `#FF0000` (red), `#00FF00` (green), `#0000FF` (blue)
- **Named Colors**: `red`, `green`, `blue`, `black`, `white`
- **RGB Values**: `rgb(255,0,0)` (red), `rgb(0,255,0)` (green)

```elixir
# Hex colors
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(spreadsheet, "Sheet1", "shape1", "#FF6B35")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_color(spreadsheet, "Sheet1", "shape1", "#004225")

# Named colors
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(spreadsheet, "Sheet1", "shape2", "lightblue")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_color(spreadsheet, "Sheet1", "shape2", "darkblue")
```

## Examples

### Example 1: Basic Shapes Collection

Create a collection of basic shapes with different colors and styles:

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Rectangle with green fill
:ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "Sheet1", "rect1")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, "Sheet1", "rect1", "rect")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_style(spreadsheet, "Sheet1", "rect1",
  "position:absolute;left:50pt;top:50pt;width:100pt;height:60pt")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_filled(spreadsheet, "Sheet1", "rect1", true)
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(spreadsheet, "Sheet1", "rect1", "#90EE90")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroked(spreadsheet, "Sheet1", "rect1", true)
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_color(spreadsheet, "Sheet1", "rect1", "#006400")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_weight(spreadsheet, "Sheet1", "rect1", "2pt")

# Circle with blue fill
:ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "Sheet1", "circle1")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, "Sheet1", "circle1", "oval")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_style(spreadsheet, "Sheet1", "circle1",
  "position:absolute;left:200pt;top:50pt;width:80pt;height:80pt")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_filled(spreadsheet, "Sheet1", "circle1", true)
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(spreadsheet, "Sheet1", "circle1", "#87CEEB")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroked(spreadsheet, "Sheet1", "circle1", true)
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_color(spreadsheet, "Sheet1", "circle1", "#4682B4")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_weight(spreadsheet, "Sheet1", "circle1", "3pt")

# Line connector
:ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "Sheet1", "line1")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, "Sheet1", "line1", "line")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_style(spreadsheet, "Sheet1", "line1",
  "position:absolute;left:150pt;top:85pt;width:50pt;height:0pt")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_filled(spreadsheet, "Sheet1", "line1", false)
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroked(spreadsheet, "Sheet1", "line1", true)
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_color(spreadsheet, "Sheet1", "line1", "#FF4500")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_weight(spreadsheet, "Sheet1", "line1", "4pt")

:ok = UmyaSpreadsheet.write(spreadsheet, "basic_shapes.xlsx")
```

### Example 2: Status Indicators

Create colored indicators for status reporting:

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add some sample data
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Task")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Status")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "Data Processing")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "Quality Check")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A4", "Final Review")

# Green indicator - Complete
:ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "Sheet1", "status_green")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, "Sheet1", "status_green", "oval")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_style(spreadsheet, "Sheet1", "status_green",
  "position:absolute;left:120pt;top:25pt;width:15pt;height:15pt")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_filled(spreadsheet, "Sheet1", "status_green", true)
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(spreadsheet, "Sheet1", "status_green", "#28a745")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroked(spreadsheet, "Sheet1", "status_green", false)

# Yellow indicator - In Progress
:ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "Sheet1", "status_yellow")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, "Sheet1", "status_yellow", "oval")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_style(spreadsheet, "Sheet1", "status_yellow",
  "position:absolute;left:120pt;top:40pt;width:15pt;height:15pt")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_filled(spreadsheet, "Sheet1", "status_yellow", true)
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(spreadsheet, "Sheet1", "status_yellow", "#ffc107")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroked(spreadsheet, "Sheet1", "status_yellow", false)

# Red indicator - Pending
:ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "Sheet1", "status_red")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, "Sheet1", "status_red", "oval")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_style(spreadsheet, "Sheet1", "status_red",
  "position:absolute;left:120pt;top:55pt;width:15pt;height:15pt")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_filled(spreadsheet, "Sheet1", "status_red", true)
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(spreadsheet, "Sheet1", "status_red", "#dc3545")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroked(spreadsheet, "Sheet1", "status_red", false)

:ok = UmyaSpreadsheet.write(spreadsheet, "status_indicators.xlsx")
```

### Example 3: Diagram with Connectors

Create a simple flow diagram using rectangles and connecting lines:

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Process box 1
:ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "Sheet1", "process1")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, "Sheet1", "process1", "rect")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_style(spreadsheet, "Sheet1", "process1",
  "position:absolute;left:50pt;top:100pt;width:120pt;height:60pt")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_filled(spreadsheet, "Sheet1", "process1", true)
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(spreadsheet, "Sheet1", "process1", "#E1F5FE")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroked(spreadsheet, "Sheet1", "process1", true)
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_color(spreadsheet, "Sheet1", "process1", "#0277BD")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_weight(spreadsheet, "Sheet1", "process1", "2pt")

# Arrow connector
:ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "Sheet1", "arrow1")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, "Sheet1", "arrow1", "line")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_style(spreadsheet, "Sheet1", "arrow1",
  "position:absolute;left:170pt;top:130pt;width:60pt;height:0pt")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_filled(spreadsheet, "Sheet1", "arrow1", false)
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroked(spreadsheet, "Sheet1", "arrow1", true)
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_color(spreadsheet, "Sheet1", "arrow1", "#424242")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_weight(spreadsheet, "Sheet1", "arrow1", "3pt")

# Process box 2
:ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "Sheet1", "process2")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, "Sheet1", "process2", "rect")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_style(spreadsheet, "Sheet1", "process2",
  "position:absolute;left:230pt;top:100pt;width:120pt;height:60pt")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_filled(spreadsheet, "Sheet1", "process2", true)
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(spreadsheet, "Sheet1", "process2", "#F3E5F5")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroked(spreadsheet, "Sheet1", "process2", true)
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_color(spreadsheet, "Sheet1", "process2", "#7B1FA2")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_weight(spreadsheet, "Sheet1", "process2", "2pt")

:ok = UmyaSpreadsheet.write(spreadsheet, "flow_diagram.xlsx")
```

### Example 4: Highlighted Data Regions

Use VML shapes to highlight important data regions:

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add sample data
for row <- 1..10 do
  UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A#{row}", "Item #{row}")
  UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B#{row}", :rand.uniform(100))
  UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C#{row}", :rand.uniform(1000))
end

# Highlight header row
:ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "Sheet1", "header_highlight")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, "Sheet1", "header_highlight", "rect")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_style(spreadsheet, "Sheet1", "header_highlight",
  "position:absolute;left:5pt;top:5pt;width:200pt;height:15pt")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_filled(spreadsheet, "Sheet1", "header_highlight", true)
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(spreadsheet, "Sheet1", "header_highlight", "#FFD700")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroked(spreadsheet, "Sheet1", "header_highlight", false)

# Highlight important data range
:ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "Sheet1", "data_highlight")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, "Sheet1", "data_highlight", "rect")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_style(spreadsheet, "Sheet1", "data_highlight",
  "position:absolute;left:5pt;top:65pt;width:200pt;height:45pt")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_filled(spreadsheet, "Sheet1", "data_highlight", true)
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(spreadsheet, "Sheet1", "data_highlight", "#FFEBEE")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroked(spreadsheet, "Sheet1", "data_highlight", true)
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_color(spreadsheet, "Sheet1", "data_highlight", "#FF5722")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_weight(spreadsheet, "Sheet1", "data_highlight", "1pt")

:ok = UmyaSpreadsheet.write(spreadsheet, "highlighted_data.xlsx")
```

## Best Practices

### Shape Naming

Use descriptive and consistent shape IDs:

```elixir
# Good: Descriptive names
"header_rectangle"
"status_indicator_green"
"connector_line_1"
"highlight_box_summary"

# Avoid: Generic names
"shape1"
"rect"
"thing"
```

### Performance Considerations

- **Batch Operations**: Group related shape operations together
- **Memory Usage**: VML shapes add to file size; use sparingly for large datasets
- **Shape Limits**: Avoid creating excessive numbers of shapes in a single worksheet

```elixir
# Efficient: Group operations for a single shape
defp create_status_indicator(spreadsheet, sheet_name, shape_id, position, color) do
  with :ok <- UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, sheet_name, shape_id),
       :ok <- UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, sheet_name, shape_id, "oval"),
       :ok <- UmyaSpreadsheet.VmlDrawing.set_shape_style(spreadsheet, sheet_name, shape_id, position),
       :ok <- UmyaSpreadsheet.VmlDrawing.set_shape_filled(spreadsheet, sheet_name, shape_id, true),
       :ok <- UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(spreadsheet, sheet_name, shape_id, color),
       :ok <- UmyaSpreadsheet.VmlDrawing.set_shape_stroked(spreadsheet, sheet_name, shape_id, false) do
    :ok
  else
    error -> error
  end
end
```

### Positioning and Layout

- **Consistent Units**: Use consistent units throughout your shapes (prefer `pt` for precision)
- **Grid Alignment**: Align shapes to a consistent grid for professional appearance
- **Relative Positioning**: Consider using relative measurements for responsive layouts

```elixir
# Consistent grid-based positioning
base_left = 50   # Base left position
base_top = 100   # Base top position
shape_width = 80
shape_spacing = 100

# Create shapes on a grid
for {shape_id, index} <- Enum.with_index(["shape1", "shape2", "shape3"]) do
  left = base_left + (index * shape_spacing)
  style = "position:absolute;left:#{left}pt;top:#{base_top}pt;width:#{shape_width}pt;height:60pt"

  :ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "Sheet1", shape_id)
  :ok = UmyaSpreadsheet.VmlDrawing.set_shape_style(spreadsheet, "Sheet1", shape_id, style)
end
```

### Color and Style Consistency

Define color palettes and style constants:

```elixir
defmodule MyApp.VmlStyles do
  # Color palette
  @primary_color "#2196F3"
  @secondary_color "#FFC107"
  @success_color "#4CAF50"
  @warning_color "#FF9800"
  @error_color "#F44336"

  # Standard dimensions
  @indicator_size "15pt"
  @box_height "60pt"
  @stroke_weight "2pt"

  def create_success_indicator(spreadsheet, sheet_name, shape_id, left, top) do
    style = "position:absolute;left:#{left}pt;top:#{top}pt;width:#{@indicator_size};height:#{@indicator_size}"

    with :ok <- UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, sheet_name, shape_id),
         :ok <- UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, sheet_name, shape_id, "oval"),
         :ok <- UmyaSpreadsheet.VmlDrawing.set_shape_style(spreadsheet, sheet_name, shape_id, style),
         :ok <- UmyaSpreadsheet.VmlDrawing.set_shape_filled(spreadsheet, sheet_name, shape_id, true),
         :ok <- UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(spreadsheet, sheet_name, shape_id, @success_color) do
      :ok
    end
  end
end
```

## Limitations

VML (Vector Markup Language) has several significant limitations that you should be aware of before implementing VML shapes in your applications.

### 🚨 **Critical Known Issues**

#### **File Reading Problems**

**CRITICAL**: Reading back files containing VML shapes may cause relationship errors in the underlying Rust library.

**Impact**: Files created with VML shapes might not be readable by the same application that created them.

**Workarounds**:

- Create VML shapes in new files rather than modifying existing ones
- Test VML functionality by writing files and verifying they can be opened in Excel
- Avoid reading back VML-containing files in the same process

#### **Shape ID Hash Collisions**

Shape IDs are converted to numeric IDs using hash-based approach, which can lead to:

- Hash collisions causing shape ID conflicts
- Inconsistent string-to-numeric conversion
- Unpredictable shape identification in complex files

### **Technical and Format Limitations**

#### **VML vs. DrawingML**

VML has several limitations compared to modern DrawingML:

- **Limited Shape Types**: Only basic geometric shapes (`rect`, `oval`, `line`, `polyline`, `roundrect`, `arc`)
- **No Custom Paths**: Cannot create complex custom vector paths
- **Basic Styling**: Limited styling options compared to modern drawing formats
- **No Text Support**: VML shapes created through this API don't support text content
- **Legacy Format**: Primarily for backward compatibility with older Excel versions

#### **Color Format Restrictions**

Limited color format support:

- Only supports `#RRGGBB` hex format and `rgb(r,g,b)` format
- No support for named colors (like "red", "blue", "green")
- No HSL, CMYK, or other color space formats
- No transparency or alpha channel support

#### **CSS Style Limitations**

- No validation for CSS-style positioning properties
- Limited support for complex positioning and sizing
- Manual CSS string construction required
- No helper functions for common positioning patterns

### **Browser and Application Compatibility**

#### **Limited Application Support**

- **Excel Versions**: Full support in Excel 2003 and later
- **LibreOffice**: Limited VML support; some shapes may not display correctly
- **Excel Online/Web**: Inconsistent support for VML features
- **Web Browsers**: VML is deprecated in modern web browsers
- **Third-party Applications**: Unpredictable rendering across different spreadsheet applications

#### **Cross-Platform Issues**

- Different behavior between Windows and macOS Excel
- Inconsistent rendering across Excel versions
- Some VML features may not render consistently

### **Performance and Memory Limitations**

#### **File Size Impact**

- VML shapes significantly increase file size
- Each shape adds substantial XML overhead
- No optimization for multiple similar shapes
- Complex shapes can lead to large file sizes

#### **Memory Usage**

Each VML shape creates multiple memory allocations:

- Boxed objects for fill, stroke, shadow, path, and text properties
- Additional overhead for shape collections and relationships
- No memory pooling or optimization for bulk operations

#### **Processing Performance**

- No batch operations for creating multiple shapes
- Must create and configure shapes individually
- No efficient way to apply same style to multiple shapes
- Linear performance degradation with shape count

### **API Design Limitations**

#### **Missing Features**

- **No Shape Positioning Helpers**: Must manually specify CSS positioning coordinates
- **No Shape Bounds Validation**: Shapes can be positioned outside visible worksheet area
- **No Shape Interaction Support**: All shapes are static drawing objects
- **No Animation or Dynamic Properties**: No support for interactive or animated elements
- **No Shape Grouping**: Cannot group shapes or manage shape collections
- **No Shape Templates**: No reusable shape configuration system

#### **Error Handling Limitations**

- Generic error messages don't provide specific failure context
- Limited validation for input parameters
- No detailed debugging information for positioning issues
- Resource locking problems in multi-threaded scenarios

### **Development and Testing Challenges**

#### **Limited Debugging Support**

- Shape visibility issues are difficult to diagnose programmatically
- No validation of positioning coordinates against worksheet bounds
- Minimal logging or error context for troubleshooting
- No programmatic way to validate shape rendering

#### **Testing Complexity**

- Must open generated files in Excel to verify visual appearance
- Cannot automate visual testing of shape rendering
- Difficult to validate cross-platform compatibility
- Manual testing required for each target Excel version

### **Architectural Limitations**

#### **Thread Safety Concerns**

- Mutex contention in multi-threaded scenarios
- Manual guard dropping required to prevent deadlocks
- No built-in concurrency support for shape operations
- Race conditions possible with shared spreadsheet resources

#### **Resource Management**

- Manual resource lifecycle management required
- No automatic cleanup of unused shapes
- Potential memory leaks with complex shape hierarchies
- Limited garbage collection optimization

### **Migration and Future-Proofing**

#### **Technology Obsolescence**

- VML is legacy technology in modern Excel
- Microsoft recommends DrawingML for new applications
- Limited future development or feature additions
- Potential removal in future Excel versions

#### **Upgrade Path Limitations**

- No automatic migration from VML to DrawingML
- Manual conversion required for shape modernization
- Breaking changes possible in future library versions
- Limited backward compatibility guarantees

## Troubleshooting

### Critical Issues and Solutions

#### **File Reading Failures**

**Problem**: Files containing VML shapes cannot be read back or cause relationship errors.

**Root Cause**: Known limitation in the underlying Rust library's VML relationship handling.

**Immediate Solutions**:

1. **Use VML for write-only operations** - Don't attempt to read files containing VML shapes
2. **Create in new files** - Only add VML shapes to newly created spreadsheets
3. **Test externally** - Verify VML shapes by opening files in Excel, not programmatically

```elixir
# Safe approach - write-only VML usage
{:ok, spreadsheet} = UmyaSpreadsheet.new()
:ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "Sheet1", "shape1")
:ok = UmyaSpreadsheet.write_file(spreadsheet, "output.xlsx")
# Don't try to read this file back in the same process
```

#### **Shape ID Collisions**

**Problem**: Multiple shapes appear to have the same ID or shapes disappear unexpectedly.

**Root Cause**: Hash-based shape ID conversion can cause collisions.

**Solutions**:

1. **Use sequential numeric IDs** when possible
2. **Keep shape IDs short and unique**
3. **Test with target shape count** to identify collision thresholds

```elixir
# Better ID strategy
shapes = for i <- 1..10 do
  shape_id = "shape_#{i}"  # Simple, sequential IDs
  :ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "Sheet1", shape_id)
  shape_id
end
```

#### **Thread Safety Issues**

**Problem**: Errors when creating VML shapes in multi-threaded applications.

**Root Cause**: Mutex contention and resource locking issues.

**Solutions**:

1. **Serialize VML operations** - Don't create shapes concurrently
2. **Use process-level locking** for VML operations
3. **Create shapes in single thread** then distribute spreadsheet

```elixir
# Safe concurrent approach
Task.async(fn ->
  # All VML operations in single process
  {:ok, spreadsheet} = UmyaSpreadsheet.new()

  Enum.each(1..100, fn i ->
    :ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "Sheet1", "shape_#{i}")
  end)

  spreadsheet
end)
|> Task.await()
```

### Common Issues

#### **Shape Not Visible**

**Problem**: Shape is created but not visible in Excel.

**Diagnostic Steps**:

1. Check positioning coordinates are within worksheet bounds
2. Verify fill/stroke settings aren't making shape invisible
3. Confirm sheet name matches exactly (case-sensitive)
4. Ensure style string is properly formatted

```elixir
# Debug visibility issues
defmodule VmlDebugger do
  def debug_shape_visibility(spreadsheet, sheet_name, shape_id) do
    # Create with explicit visibility settings
    :ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, sheet_name, shape_id)

    # Set very visible properties
    style = "position:absolute;left:10pt;top:10pt;width:100pt;height:50pt;z-index:1"
    :ok = UmyaSpreadsheet.VmlDrawing.set_shape_style(spreadsheet, sheet_name, shape_id, style)

    # Bright red fill, thick black border
    :ok = UmyaSpreadsheet.VmlDrawing.set_shape_filled(spreadsheet, sheet_name, shape_id, true)
    :ok = UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(spreadsheet, sheet_name, shape_id, "#FF0000")
    :ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroked(spreadsheet, sheet_name, shape_id, true)
    :ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_color(spreadsheet, sheet_name, shape_id, "#000000")
    :ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_weight(spreadsheet, sheet_name, shape_id, "3pt")

    IO.puts("Debug shape '#{shape_id}' created with maximum visibility")
  end
end
```

#### **Color Format Errors**

**Problem**: Invalid color format errors when setting fill or stroke colors.

**Root Cause**: Limited color format validation only supports hex and rgb() formats.

**Solutions**:

1. **Use hex format**: `#RRGGBB` (e.g., `#FF0000` for red)
2. **Use rgb format**: `rgb(r,g,b)` (e.g., `rgb(255,0,0)` for red)
3. **Convert named colors** to hex before using

```elixir
defmodule ColorHelper do
  # Convert common named colors to hex
  def named_to_hex(color) do
    case String.downcase(color) do
      "red" -> "#FF0000"
      "green" -> "#00FF00"
      "blue" -> "#0000FF"
      "black" -> "#000000"
      "white" -> "#FFFFFF"
      "yellow" -> "#FFFF00"
      "orange" -> "#FFA500"
      "purple" -> "#800080"
      hex when binary_part(hex, 0, 1) == "#" -> hex
      _ -> "#000000"  # Default to black for unknown colors
    end
  end

  def set_safe_fill_color(spreadsheet, sheet_name, shape_id, color) do
    hex_color = named_to_hex(color)
    UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(spreadsheet, sheet_name, shape_id, hex_color)
  end
end
```

#### **Shape Positioning Issues**

**Problem**: Shapes appear in unexpected locations or outside worksheet bounds.

**Root Cause**: No validation of positioning coordinates or CSS style string format.

**Solutions**:

1. **Use helper functions** for reliable positioning
2. **Validate coordinates** before applying styles
3. **Use consistent units** (prefer `pt` over `px` or `em`)

```elixir
defmodule PositionHelper do
  def create_positioned_shape(spreadsheet, sheet_name, shape_id, opts) do
    left = Keyword.get(opts, :left, 10)
    top = Keyword.get(opts, :top, 10)
    width = Keyword.get(opts, :width, 100)
    height = Keyword.get(opts, :height, 50)

    # Validate coordinates are reasonable
    unless left >= 0 && top >= 0 && width > 0 && height > 0 do
      raise ArgumentError, "Invalid positioning coordinates"
    end

    # Create shape
    :ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, sheet_name, shape_id)

    # Apply validated positioning
    style = "position:absolute;left:#{left}pt;top:#{top}pt;width:#{width}pt;height:#{height}pt"
    :ok = UmyaSpreadsheet.VmlDrawing.set_shape_style(spreadsheet, sheet_name, shape_id, style)

    {:ok, %{left: left, top: top, width: width, height: height}}
  end

  def position_in_cell(spreadsheet, sheet_name, shape_id, cell_ref) do
    # Rough cell positioning (Excel default cell size ~64pt x 20pt)
    {col, row} = parse_cell_reference(cell_ref)
    left = (col - 1) * 64
    top = (row - 1) * 20

    create_positioned_shape(spreadsheet, sheet_name, shape_id,
                          left: left, top: top, width: 60, height: 18)
  end

  defp parse_cell_reference(cell_ref) do
    # Simple A1 reference parser (extend as needed)
    # This is a basic implementation
    {1, 1}  # Default to A1 for now
  end
end
```

### Error Messages and Solutions

#### **"Failed to lock spreadsheet resource"**

**Cause**: Thread contention or deadlock in multi-threaded environment.

**Solutions**:

1. Ensure VML operations are serialized
2. Check for competing access to the same spreadsheet resource
3. Implement retry logic with exponential backoff

```elixir
defmodule SafeVmlOperations do
  def with_retry(operation, max_retries \\ 3) do
    do_with_retry(operation, max_retries, 100)
  end

  defp do_with_retry(operation, 0, _delay), do: {:error, :max_retries_exceeded}

  defp do_with_retry(operation, retries_left, delay) do
    case operation.() do
      :ok -> :ok
      {:error, reason} when reason =~ "Failed to lock" ->
        Process.sleep(delay)
        do_with_retry(operation, retries_left - 1, delay * 2)
      {:error, reason} -> {:error, reason}
    end
  end
end

# Usage
SafeVmlOperations.with_retry(fn ->
  UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "Sheet1", "shape1")
end)
```

#### **"Invalid shape type"**

**Cause**: Using unsupported shape type string.

**Valid Types**: `"rect"`, `"oval"`, `"line"`, `"polyline"`, `"roundrect"`, `"arc"`

```elixir
defmodule ShapeTypeValidator do
  @valid_types ["rect", "oval", "line", "polyline", "roundrect", "arc"]

  def validate_shape_type(shape_type) do
    if shape_type in @valid_types do
      {:ok, shape_type}
    else
      {:error, "Invalid shape type '#{shape_type}'. Valid types: #{inspect(@valid_types)}"}
    end
  end

  def set_safe_shape_type(spreadsheet, sheet_name, shape_id, shape_type) do
    case validate_shape_type(shape_type) do
      {:ok, valid_type} ->
        UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, sheet_name, shape_id, valid_type)
      {:error, reason} ->
        {:error, reason}
    end
  end
end
```

### Performance Optimization

#### **Managing Large Numbers of Shapes**

**Problem**: Performance degrades significantly with many VML shapes.

**Solutions**:

1. **Batch operations** where possible
2. **Limit shape count** per worksheet
3. **Consider alternatives** for large datasets

```elixir
defmodule VmlPerformance do
  @max_shapes_per_sheet 50  # Recommended limit

  def create_shapes_efficiently(spreadsheet, sheet_name, shape_configs) do
    if length(shape_configs) > @max_shapes_per_sheet do
      IO.warn("Creating #{length(shape_configs)} shapes may impact performance. Consider alternatives.")
    end

    Enum.with_index(shape_configs, 1)
    |> Enum.map(fn {config, index} ->
      shape_id = "shape_#{index}"

      # Create shape
      :ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, sheet_name, shape_id)

      # Apply configuration efficiently
      apply_shape_config(spreadsheet, sheet_name, shape_id, config)

      shape_id
    end)
  end

  defp apply_shape_config(spreadsheet, sheet_name, shape_id, config) do
    # Apply all properties in sequence to minimize NIF calls
    if config[:type] do
      :ok = UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, sheet_name, shape_id, config[:type])
    end

    if config[:style] do
      :ok = UmyaSpreadsheet.VmlDrawing.set_shape_style(spreadsheet, sheet_name, shape_id, config[:style])
    end

    # ... apply other properties
  end
end
```

### Development Best Practices

#### **Testing VML Implementation**

```elixir
defmodule VmlTesting do
  def test_vml_functionality() do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Test basic shape creation
    assert :ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "Sheet1", "test_shape")

    # Test all property setters
    assert :ok = UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, "Sheet1", "test_shape", "rect")
    assert :ok = UmyaSpreadsheet.VmlDrawing.set_shape_filled(spreadsheet, "Sheet1", "test_shape", true)
    assert :ok = UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(spreadsheet, "Sheet1", "test_shape", "#FF0000")

    # Write to temporary file for manual verification
    test_file = "/tmp/vml_test_#{System.system_time(:millisecond)}.xlsx"
    :ok = UmyaSpreadsheet.write_file(spreadsheet, test_file)

    IO.puts("VML test file created: #{test_file}")
    IO.puts("Open in Excel to verify VML shapes are visible")

    test_file
  end
end
```

### Alternative Approaches

When VML limitations are too restrictive, consider these alternatives:

#### **Conditional Formatting for Simple Indicators**

```elixir
# Instead of VML shapes for status indicators
UmyaSpreadsheet.ConditionalFormatting.add_cell_value_rule(
  spreadsheet, "Sheet1", "A1:A10",
  operator: :equal,
  value: "complete",
  format: %{background_color: "#00FF00"}
)
```

#### **Unicode Symbols for Simple Graphics**

```elixir
# Use Unicode symbols instead of VML shapes
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "●")  # Bullet
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "▲")  # Triangle
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "■")  # Square
```

### Getting Help

If you encounter issues with VML drawings:

1. **Check Critical Issues** section above for known problems
2. **Review limitations** to ensure VML is appropriate for your use case
3. **Test with minimal examples** to isolate issues
4. **Consider alternatives** when VML limitations are blocking
5. **Verify Excel compatibility** by opening generated files manually

For general debugging:

1. Check the [Limitations Guide](limitations.html) for known constraints
2. Review the [Troubleshooting Guide](troubleshooting.html) for general debugging tips
3. Verify your VML usage follows the examples in this guide

---

*This guide covers the VML Drawing functionality available in UmyaSpreadsheet. For more information about other drawing features, see the [Shapes and Drawing Guide](shapes_and_drawing.html).*
