# Helper Modules

This directory contains helper modules that provide common functionality and abstractions for the Umya Spreadsheet Elixir wrapper.

## Available Helpers

### 1. Color Helper (`color_helper.rs`)

Handles color parsing and creation:

- `parse_color(color_str: &str) -> Result<String, Atom>`: Parses color strings (named colors or hex values) to ARGB format.
- `create_color_object(color_str: &str) -> Result<Color, Atom>`: Creates a `Color` object from a color string.

### 2. Path Helper (`path_helper.rs`)

Handles file path validation:

- `find_valid_file_path(file_path: &str) -> Option<String>`: Checks if the provided path exists on the filesystem.

### 3. Alignment Helper (`alignment_helper.rs`)

Handles cell alignment:

- `parse_horizontal_alignment(alignment: &str) -> HorizontalAlignmentValues`: Converts a string to a horizontal alignment enum.
- `parse_vertical_alignment(alignment: &str) -> VerticalAlignmentValues`: Converts a string to a vertical alignment enum.

### 4. Format Helper (`format_helper.rs`)

Handles CSV formatting options:

- `parse_csv_encoding(encoding: &str) -> CsvEncodeValues`: Converts an encoding string to the corresponding enum.
- `create_csv_writer_options(encoding: &str, do_trim: bool, wrap_with_char: &str) -> CsvWriterOption`: Creates a CSV writer options object.

### 5. Style Helper (`style_helper.rs`)

Handles cell styling operations:

- `create_pattern_fill(color: Color) -> PatternFill`: Creates a pattern fill with the specified foreground color.
- `apply_cell_style(sheet: &mut Worksheet, cell_address: &str, bg_color: Option<Color>, font_color: Option<Color>, font_size: Option<f64>, is_bold: Option<bool>)`: Sets multiple style properties on a cell in a single operation.
- `apply_row_style(sheet: &mut Worksheet, row_number: u32, bg_color: Option<Color>, font_color: Option<Color>, font_size: Option<f64>, is_bold: Option<bool>)`: Sets multiple style properties on an entire row in a single operation.
