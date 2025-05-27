# Rich Text Formatting

UmyaSpreadsheet provides comprehensive Rich Text support, allowing you to create cells with multiple formatting styles within a single cell. This enables you to create professional-looking spreadsheets with complex text formatting that goes beyond simple cell-level styling.

## Overview

Rich Text in Excel allows you to format different parts of a cell's text with different styles. For example, you can have **bold** text followed by *italic* text, or text in different colors and sizes, all within the same cell.

Key features of UmyaSpreadsheet's Rich Text support:

- **Multiple font styles** - Bold, italic, underline, strikethrough
- **Custom colors** - Hex colors and named colors for text
- **Font customization** - Different font names and sizes within the same cell
- **Mixed formatting** - Combine multiple formatting options
- **HTML support** - Create rich text from HTML strings
- **Idiomatic Elixir API** - Uses atom keys (`:bold`, `:italic`) for easy integration

## Quick Start

### Creating Basic Rich Text

```elixir
# Load or create a spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Create a rich text object
{:ok, rich_text} = UmyaSpreadsheet.RichText.create()

# Add formatted text segments
{:ok, rich_text} = UmyaSpreadsheet.RichText.add_formatted_text(
  rich_text,
  "Hello ",
  %{bold: true, color: "FF0000"}
)

{:ok, rich_text} = UmyaSpreadsheet.RichText.add_formatted_text(
  rich_text,
  "World!",
  %{italic: true, color: "0000FF"}
)

# Apply to a cell
{:ok, spreadsheet} = UmyaSpreadsheet.RichText.set_cell_rich_text(
  spreadsheet,
  "Sheet1",
  "A1",
  rich_text
)
```

### Creating Rich Text from HTML

```elixir
# Create rich text from HTML string
{:ok, rich_text} = UmyaSpreadsheet.RichText.create_from_html(
  "<b>Bold text</b> and <i style='color: blue;'>blue italic text</i>"
)

# Apply to cell
{:ok, spreadsheet} = UmyaSpreadsheet.RichText.set_cell_rich_text(
  spreadsheet,
  "Sheet1",
  "B1",
  rich_text
)
```

## Core Functions

### Creating Rich Text Objects

#### `create/0`

Creates a new empty rich text object:

```elixir
{:ok, rich_text} = UmyaSpreadsheet.RichText.create()
```

**Returns**: `{:ok, rich_text}` or `{:error, reason}`

#### `create_from_html/1`

Creates rich text from an HTML string:

```elixir
html = "<b>Bold</b> and <i>italic</i> text"
{:ok, rich_text} = UmyaSpreadsheet.RichText.create_from_html(html)
```

**Parameters**:

- `html_string` - HTML string with formatting tags

**Returns**: `{:ok, rich_text}` or `{:error, reason}`

**Supported HTML Tags**:

- `<b>` or `<strong>` - Bold text
- `<i>` or `<em>` - Italic text
- `<u>` - Underlined text
- `<s>` or `<strike>` - Strikethrough text
- `<span style="color: #FF0000;">` - Colored text
- `<span style="font-size: 14px;">` - Custom font size
- `<span style="font-family: Arial;">` - Custom font family

### Managing Text Elements

#### `add_formatted_text/3`

Adds formatted text to a rich text object:

```elixir
{:ok, rich_text} = UmyaSpreadsheet.RichText.add_formatted_text(
  rich_text,
  "Important: ",
  %{
    bold: true,
    color: "FF0000",    # Red color in hex
    font_size: 14,
    font_name: "Arial"
  }
)
```

**Parameters**:

- `rich_text` - Existing rich text object
- `text` - Text content to add
- `font_properties` - Map of formatting options (see Font Properties section)

**Returns**: `{:ok, updated_rich_text}` or `{:error, reason}`

#### `create_text_element/2`

Creates a standalone text element:

```elixir
{:ok, element} = UmyaSpreadsheet.RichText.create_text_element(
  "Bold text",
  %{bold: true}
)
```

#### `add_text_element/2`

Adds an existing text element to rich text:

```elixir
{:ok, rich_text} = UmyaSpreadsheet.RichText.add_text_element(rich_text, element)
```

### Cell Operations

#### `set_cell_rich_text/4`

Applies rich text formatting to a specific cell:

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.RichText.set_cell_rich_text(
  spreadsheet,
  "Sheet1",    # Worksheet name
  "A1",        # Cell reference
  rich_text    # Rich text object
)
```

**Parameters**:

- `spreadsheet` - Spreadsheet object
- `worksheet_name` - Name of the target worksheet
- `cell_reference` - Cell reference (e.g., "A1", "B2")
- `rich_text` - Rich text object to apply

**Returns**: `{:ok, updated_spreadsheet}` or `{:error, reason}`

#### `get_cell_rich_text/3`

Retrieves rich text from a cell:

```elixir
{:ok, rich_text} = UmyaSpreadsheet.RichText.get_cell_rich_text(
  spreadsheet,
  "Sheet1",
  "A1"
)
```

**Returns**: `{:ok, rich_text}` or `{:error, reason}`

### Inspection and Export

#### `get_elements/1`

Gets all text elements from a rich text object:

```elixir
{:ok, elements} = UmyaSpreadsheet.RichText.get_elements(rich_text)
```

#### `get_element_text/1`

Extracts text content from a text element:

```elixir
{:ok, text} = UmyaSpreadsheet.RichText.get_element_text(element)
```

#### `get_element_font_properties/1`

Gets font properties from a text element:

```elixir
{:ok, properties} = UmyaSpreadsheet.RichText.get_element_font_properties(element)
# Returns: %{bold: true, italic: false, color: "FF0000", ...}
```

#### `to_html/1`

Converts rich text to HTML string:

```elixir
{:ok, html_string} = UmyaSpreadsheet.RichText.to_html(rich_text)
# Returns: "<b>Bold text</b> and <i>italic text</i>"
```

## Font Properties

Font properties are specified using maps with atom keys. All properties are optional:

### Basic Text Styling

```elixir
%{
  bold: true,           # Bold text
  italic: true,         # Italic text
  underline: true,      # Underlined text
  strikethrough: true   # Strikethrough text
}
```

### Font Appearance

```elixir
%{
  font_name: "Arial",   # Font family name
  font_size: 14,        # Font size in points
  color: "FF0000"       # Text color in hex (without #)
}
```

### Color Formats

Colors can be specified in several formats:

```elixir
# Hex color (preferred)
%{color: "FF0000"}      # Red
%{color: "00FF00"}      # Green
%{color: "0000FF"}      # Blue

# Named colors (basic support)
%{color: "red"}
%{color: "blue"}
%{color: "green"}
```

### Complete Example

```elixir
font_props = %{
  bold: true,
  italic: false,
  underline: true,
  strikethrough: false,
  font_name: "Calibri",
  font_size: 12,
  color: "336699"       # Dark blue
}
```

## Practical Examples

### 1. Creating a Formatted Header

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()
{:ok, rich_text} = UmyaSpreadsheet.RichText.create()

# Company name in large, bold text
{:ok, rich_text} = UmyaSpreadsheet.RichText.add_formatted_text(
  rich_text,
  "ACME Corporation\n",
  %{bold: true, font_size: 16, color: "2E4057"}
)

# Subtitle in smaller, italic text
{:ok, rich_text} = UmyaSpreadsheet.RichText.add_formatted_text(
  rich_text,
  "Financial Report - Q4 2024",
  %{italic: true, font_size: 12, color: "5A6B7D"}
)

{:ok, spreadsheet} = UmyaSpreadsheet.RichText.set_cell_rich_text(
  spreadsheet, "Sheet1", "A1", rich_text
)
```

### 2. Status Indicators with Colors

```elixir
defmodule StatusFormatter do
  def format_status(status) do
    {:ok, rich_text} = UmyaSpreadsheet.RichText.create()

    {status_text, color} = case status do
      :success -> {"✓ COMPLETED", "008000"}    # Green
      :warning -> {"⚠ IN PROGRESS", "FFA500"}  # Orange
      :error   -> {"✗ FAILED", "FF0000"}       # Red
    end

    UmyaSpreadsheet.RichText.add_formatted_text(
      rich_text,
      status_text,
      %{bold: true, color: color}
    )
  end
end

# Usage
{:ok, rich_text} = StatusFormatter.format_status(:success)
{:ok, spreadsheet} = UmyaSpreadsheet.RichText.set_cell_rich_text(
  spreadsheet, "Sheet1", "B1", rich_text
)
```

### 3. Multi-line Instructions

```elixir
{:ok, rich_text} = UmyaSpreadsheet.RichText.create()

# Step header
{:ok, rich_text} = UmyaSpreadsheet.RichText.add_formatted_text(
  rich_text,
  "Setup Instructions:\n",
  %{bold: true, underline: true, font_size: 14}
)

# Step 1
{:ok, rich_text} = UmyaSpreadsheet.RichText.add_formatted_text(
  rich_text,
  "1. ",
  %{bold: true, color: "0066CC"}
)

{:ok, rich_text} = UmyaSpreadsheet.RichText.add_formatted_text(
  rich_text,
  "Download the installer\n",
  %{}
)

# Step 2 with emphasis
{:ok, rich_text} = UmyaSpreadsheet.RichText.add_formatted_text(
  rich_text,
  "2. ",
  %{bold: true, color: "0066CC"}
)

{:ok, rich_text} = UmyaSpreadsheet.RichText.add_formatted_text(
  rich_text,
  "Run as ",
  %{}
)

{:ok, rich_text} = UmyaSpreadsheet.RichText.add_formatted_text(
  rich_text,
  "Administrator",
  %{bold: true, italic: true, color: "CC0000"}
)
```

### 4. HTML Import for Complex Formatting

```elixir
html_content = """
<b style="font-size: 16px; color: #2C3E50;">Product Overview</b><br/>
<i>Model: </i><b>ABC-123</b><br/>
<span style="color: #27AE60;">✓ In Stock</span><br/>
<span style="font-size: 10px; color: #7F8C8D;">Last updated: #{Date.utc_today()}</span>
"""

{:ok, rich_text} = UmyaSpreadsheet.RichText.create_from_html(html_content)
{:ok, spreadsheet} = UmyaSpreadsheet.RichText.set_cell_rich_text(
  spreadsheet, "Sheet1", "C1", rich_text
)
```

## Error Handling

Rich Text functions return `{:ok, result}` or `{:error, reason}` tuples:

```elixir
case UmyaSpreadsheet.RichText.create_from_html(html_string) do
  {:ok, rich_text} ->
    # Success - use rich_text
    IO.puts("Rich text created successfully")

  {:error, reason} ->
    # Handle error
    IO.puts("Failed to create rich text: #{reason}")
end
```

Common error scenarios:

- Invalid HTML in `create_from_html/1`
- Invalid font properties in formatting functions
- Cell reference errors in cell operations
- Resource allocation failures

## Best Practices

### 1. Performance Considerations

- **Reuse rich text objects** when applying the same formatting to multiple cells
- **Batch operations** when setting multiple rich text cells
- **Limit complexity** - very complex rich text can impact file size and performance

```elixir
# Good: Reuse rich text object
{:ok, header_format} = UmyaSpreadsheet.RichText.create_from_html("<b>Header</b>")

for col <- ["A", "B", "C"] do
  UmyaSpreadsheet.RichText.set_cell_rich_text(
    spreadsheet, "Sheet1", "#{col}1", header_format
  )
end
```

### 2. Color Consistency

- **Use a color palette** - define your colors as constants
- **Consider accessibility** - ensure sufficient contrast
- **Test in Excel** - verify colors display correctly

```elixir
defmodule ColorPalette do
  def primary(), do: "2C3E50"      # Dark blue-gray
  def success(), do: "27AE60"      # Green
  def warning(), do: "F39C12"      # Orange
  def danger(), do: "E74C3C"       # Red
  def muted(), do: "95A5A6"        # Light gray
end
```

### 3. Font Selection

- **Use system fonts** - ensure compatibility across platforms
- **Consistent sizing** - maintain visual hierarchy
- **Test on target platforms** - verify font availability

```elixir
# Safe font choices
font_stack = ["Calibri", "Arial", "Helvetica", "sans-serif"]
primary_font = "Calibri"  # Default Excel font
```

### 4. HTML Import Guidelines

- **Validate HTML** before importing
- **Use supported tags** only
- **Test complex formatting** to ensure correct conversion
- **Fallback to manual creation** for unsupported features

```elixir
defmodule HtmlValidator do
  @supported_tags ~w[b strong i em u s strike span br]

  def validate_html(html) do
    # Basic validation logic
    # Return {:ok, html} or {:error, reason}
  end
end
```

## Troubleshooting

### Common Issues

1. **Colors not displaying correctly**
   - Ensure hex colors don't include the `#` symbol
   - Verify 6-digit hex format (e.g., "FF0000" not "F00")

2. **Font not appearing**
   - Check font name spelling and availability
   - Use quoted font names for fonts with spaces

3. **HTML import failing**
   - Validate HTML syntax
   - Use only supported HTML tags
   - Check for proper tag closure

4. **Rich text not saving**
   - Ensure proper error handling
   - Check cell reference format
   - Verify worksheet exists

### Debug Techniques

```elixir
# Inspect rich text content
{:ok, elements} = UmyaSpreadsheet.RichText.get_elements(rich_text)
Enum.each(elements, fn element ->
  {:ok, text} = UmyaSpreadsheet.RichText.get_element_text(element)
  {:ok, props} = UmyaSpreadsheet.RichText.get_element_font_properties(element)
  IO.inspect({text, props})
end)

# Convert to HTML for debugging
{:ok, html} = UmyaSpreadsheet.RichText.to_html(rich_text)
IO.puts("Generated HTML: #{html}")
```

## Integration with Other Features

Rich Text works seamlessly with other UmyaSpreadsheet features:

### With Cell Styling

```elixir
# Set cell background color
UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "A1", "F0F0F0")

# Add rich text content
{:ok, rich_text} = UmyaSpreadsheet.RichText.create_from_html("<b>Header</b>")
UmyaSpreadsheet.RichText.set_cell_rich_text(spreadsheet, "Sheet1", "A1", rich_text)
```

### With Data Validation

```elixir
# Rich text can be used in cells with data validation
UmyaSpreadsheet.add_data_validation_list(spreadsheet, "Sheet1", "B1", ["Option 1", "Option 2"])

# Then add rich text instructions
{:ok, instructions} = UmyaSpreadsheet.RichText.create_from_html(
  "<i>Select from dropdown:</i>"
)
UmyaSpreadsheet.RichText.set_cell_rich_text(spreadsheet, "Sheet1", "A1", instructions)
```

### With Hyperlinks

```elixir
# Add hyperlink to cell
UmyaSpreadsheet.add_hyperlink(spreadsheet, "Sheet1", "A1",
  "https://example.com", "Visit Website")

# Note: Rich text and hyperlinks are mutually exclusive in the same cell
# Use rich text for display text in adjacent cells
```

## Conclusion

Rich Text formatting in UmyaSpreadsheet provides powerful tools for creating professional, visually appealing spreadsheets. By combining different formatting options, colors, and fonts, you can create clear, readable documents that effectively communicate information.

For more advanced cell formatting options, see the [Advanced Cell Formatting](advanced_cell_formatting.html) guide. For basic styling features, refer to the [Styling and Formatting](styling_and_formatting.html) guide.
