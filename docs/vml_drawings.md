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

### Shape Property Retrieval

UmyaSpreadsheet provides functions to retrieve properties of VML shapes:

#### `get_shape_style/3`

Gets the CSS style of a VML shape.

```elixir
@spec get_shape_style(Spreadsheet.t(), String.t(), String.t()) ::
        {:ok, String.t()} | {:error, String.t()}
```

**Parameters:**

- `spreadsheet` - A spreadsheet struct
- `sheet_name` - The name of the worksheet
- `shape_id` - Unique identifier for the shape

**Returns:**

- `{:ok, style}` on success where style is a CSS style string
- `{:error, reason}` on failure

**Example:**

```elixir
{:ok, style} = UmyaSpreadsheet.VmlDrawing.get_shape_style(spreadsheet, "Sheet1", "myShape")
# style might be "position:absolute;left:100pt;top:100pt;width:200pt;height:100pt"
```

#### `get_shape_type/3`

Gets the type of a VML shape.

```elixir
@spec get_shape_type(Spreadsheet.t(), String.t(), String.t()) ::
        {:ok, String.t()} | {:error, String.t()}
```

**Parameters:**

- `spreadsheet` - A spreadsheet struct
- `sheet_name` - The name of the worksheet
- `shape_id` - Unique identifier for the shape

**Returns:**

- `{:ok, shape_type}` on success where shape_type is the type (e.g. "rect", "oval", etc.)
- `{:error, reason}` on failure

**Example:**

```elixir
{:ok, shape_type} = UmyaSpreadsheet.VmlDrawing.get_shape_type(spreadsheet, "Sheet1", "myShape")
# shape_type might be "rect" or "oval"
```

#### `get_shape_filled/3`

Gets whether a VML shape is filled.

```elixir
@spec get_shape_filled(Spreadsheet.t(), String.t(), String.t()) ::
        {:ok, boolean()} | {:error, String.t()}
```

**Parameters:**

- `spreadsheet` - A spreadsheet struct
- `sheet_name` - The name of the worksheet
- `shape_id` - Unique identifier for the shape

**Returns:**

- `{:ok, filled}` on success where filled is a boolean
- `{:error, reason}` on failure

**Example:**

```elixir
{:ok, is_filled} = UmyaSpreadsheet.VmlDrawing.get_shape_filled(spreadsheet, "Sheet1", "myShape")
```

#### `get_shape_fill_color/3`

Gets the fill color of a VML shape.

```elixir
@spec get_shape_fill_color(Spreadsheet.t(), String.t(), String.t()) ::
        {:ok, String.t()} | {:error, String.t()}
```

**Parameters:**

- `spreadsheet` - A spreadsheet struct
- `sheet_name` - The name of the worksheet
- `shape_id` - Unique identifier for the shape

**Returns:**

- `{:ok, fill_color}` on success where fill_color is a color string (e.g. "#FF0000")
- `{:error, reason}` on failure

**Example:**

```elixir
{:ok, color} = UmyaSpreadsheet.VmlDrawing.get_shape_fill_color(spreadsheet, "Sheet1", "myShape")
# color might be "#CCFFCC"
```

#### `get_shape_stroked/3`

Gets whether a VML shape has a stroke (outline).

```elixir
@spec get_shape_stroked(Spreadsheet.t(), String.t(), String.t()) ::
        {:ok, boolean()} | {:error, String.t()}
```

**Parameters:**

- `spreadsheet` - A spreadsheet struct
- `sheet_name` - The name of the worksheet
- `shape_id` - Unique identifier for the shape

**Returns:**

- `{:ok, stroked}` on success where stroked is a boolean
- `{:error, reason}` on failure

**Example:**

```elixir
{:ok, has_stroke} = UmyaSpreadsheet.VmlDrawing.get_shape_stroked(spreadsheet, "Sheet1", "myShape")
```

#### `get_shape_stroke_color/3`

Gets the stroke (outline) color of a VML shape.

```elixir
@spec get_shape_stroke_color(Spreadsheet.t(), String.t(), String.t()) ::
        {:ok, String.t()} | {:error, String.t()}
```

**Parameters:**

- `spreadsheet` - A spreadsheet struct
- `sheet_name` - The name of the worksheet
- `shape_id` - Unique identifier for the shape

**Returns:**

- `{:ok, stroke_color}` on success where stroke_color is a color string (e.g. "#0000FF")
- `{:error, reason}` on failure

**Example:**

```elixir
{:ok, color} = UmyaSpreadsheet.VmlDrawing.get_shape_stroke_color(spreadsheet, "Sheet1", "myShape")
# color might be "#009900"
```

#### `get_shape_stroke_weight/3`

Gets the stroke (outline) weight/thickness of a VML shape.

```elixir
@spec get_shape_stroke_weight(Spreadsheet.t(), String.t(), String.t()) ::
        {:ok, String.t()} | {:error, String.t()}
```

**Parameters:**

- `spreadsheet` - A spreadsheet struct
- `sheet_name` - The name of the worksheet
- `shape_id` - Unique identifier for the shape

**Returns:**

- `{:ok, stroke_weight}` on success where stroke_weight is a string (e.g. "2pt")
- `{:error, reason}` on failure

**Example:**

```elixir
{:ok, weight} = UmyaSpreadsheet.VmlDrawing.get_shape_stroke_weight(spreadsheet, "Sheet1", "myShape")
# weight might be "2pt"
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

### Setting Fill Properties

Fill properties control the interior appearance of shapes:

```elixir
# Enable fill and set color
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_filled(spreadsheet, "Sheet1", "shape1", true)
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(spreadsheet, "Sheet1", "shape1", "#FFEECC")

# Disable fill for transparent shape
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_filled(spreadsheet, "Sheet1", "shape2", false)
```

### Setting Stroke Properties

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

### Example 4: Reading Shape Properties

The following example demonstrates how to create shapes and then read their properties:

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Create and configure a shape
:ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "Sheet1", "test_shape")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, "Sheet1", "test_shape", "oval")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_style(spreadsheet, "Sheet1", "test_shape",
  "position:absolute;left:150pt;top:150pt;width:120pt;height:80pt")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_filled(spreadsheet, "Sheet1", "test_shape", true)
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(spreadsheet, "Sheet1", "test_shape", "#33AA33")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroked(spreadsheet, "Sheet1", "test_shape", true)
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_color(spreadsheet, "Sheet1", "test_shape", "#AA3333")
:ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_weight(spreadsheet, "Sheet1", "test_shape", "2.5pt")

# Now read back the shape properties
{:ok, shape_type} = UmyaSpreadsheet.VmlDrawing.get_shape_type(spreadsheet, "Sheet1", "test_shape")
IO.puts("Shape type: #{shape_type}")  # Should print "oval"

{:ok, style} = UmyaSpreadsheet.VmlDrawing.get_shape_style(spreadsheet, "Sheet1", "test_shape")
IO.puts("Style: #{style}")  # Should print the style string with position and dimensions

{:ok, is_filled} = UmyaSpreadsheet.VmlDrawing.get_shape_filled(spreadsheet, "Sheet1", "test_shape")
IO.puts("Is filled: #{is_filled}")  # Should print "true"

{:ok, fill_color} = UmyaSpreadsheet.VmlDrawing.get_shape_fill_color(spreadsheet, "Sheet1", "test_shape")
IO.puts("Fill color: #{fill_color}")  # Should print "#33AA33"

{:ok, is_stroked} = UmyaSpreadsheet.VmlDrawing.get_shape_stroked(spreadsheet, "Sheet1", "test_shape")
IO.puts("Has stroke: #{is_stroked}")  # Should print "true"

{:ok, stroke_color} = UmyaSpreadsheet.VmlDrawing.get_shape_stroke_color(spreadsheet, "Sheet1", "test_shape")
IO.puts("Stroke color: #{stroke_color}")  # Should print "#AA3333"

{:ok, stroke_weight} = UmyaSpreadsheet.VmlDrawing.get_shape_stroke_weight(spreadsheet, "Sheet1", "test_shape")
IO.puts("Stroke weight: #{stroke_weight}")  # Should print "2.5pt"
```

This example shows the complementary nature of the setter and getter functions, allowing you to not only create and configure VML shapes but also to inspect and retrieve their properties.
