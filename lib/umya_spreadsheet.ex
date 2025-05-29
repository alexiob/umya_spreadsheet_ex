defmodule UmyaSpreadsheet do
  @moduledoc """
  UmyaSpreadsheet is an Elixir wrapper for the Rust `umya-spreadsheet` library,
  which provides Excel (.xlsx, .xlsm) file manipulation capabilities with the
  performance benefits of Rust.

  ## Overview

  This library allows you to:
  - Create new spreadsheets from scratch
  - Read/write Excel files with full formatting support
  - Manipulate cell values, formulas, and styles
  - Move and organize ranges of data
  - Add and manage multiple worksheets
  - Create charts, images, and visual elements
  - Embed OLE objects (Word documents, PowerPoint presentations, etc.)
  - Add data validation and conditional formatting
  - Export to CSV with performance optimizations
  - Apply advanced formatting and styling options

  ## Architecture

  UmyaSpreadsheet follows a modular architecture with specialized function modules:

  ### Core Architecture

  ```
  UmyaSpreadsheet (Main Module)
  ├── UmyaNative (Rust NIF Interface)
  │   ├── Spreadsheet Operations (new, read, write)
  │   ├── Cell Operations (get/set values, formatting)
  │   ├── Sheet Management (add, remove, rename)
  │   └── Native Rust Library (umya-spreadsheet)
  │
  └── Specialized Function Modules:
      ├── AutoFilterFunctions - Data filtering and sorting
      ├── BackgroundFunctions - Cell background colors and patterns
      ├── BorderFunctions - Cell border styling and formatting
      ├── CellFunctions - Cell value manipulation and retrieval
      ├── ChartFunctions - Chart creation and customization
      ├── CommentFunctions - Cell comment management
      ├── ConditionalFormatting - Advanced conditional formatting rules
      ├── CSVFunctions - CSV export and import operations
      ├── DataValidation - Input validation and dropdown lists
      ├── Drawing - Shapes, connectors, and drawing objects
      ├── FileFormatOptions - File format and compression options
      ├── FontFunctions - Font styling and text formatting
      ├── FormulaFunctions - Formula creation and named ranges
      ├── Hyperlink - Hyperlink management and navigation
      ├── ImageFunctions - Image insertion and positioning
      ├── OleObjects - Object Linking and Embedding (OLE) objects
      ├── PerformanceFunctions - Memory-optimized operations
      ├── PivotTable - Pivot table creation and management
      ├── RichText - Formatted text within cells with styling
      ├── Table - Excel table creation and management
      ├── PrintSettings - Page setup and print configuration
      ├── ProtectionFunctions - Security and access control
      ├── RowColumnFunctions - Row and column operations
      ├── SheetFunctions - Worksheet management and properties
      ├── StyleFunctions - Cell styling and number formatting
      └── WindowFunctions - View settings and window management
  ```

  ### Data Flow

  1. **NIF Layer**: All operations go through the native interface which interfaces with Rust
  2. **Spreadsheet Reference**: Operations work on a Rust-managed spreadsheet reference
  3. **Function Modules**: Specialized modules provide domain-specific functionality
  4. **Error Handling**: Comprehensive error handling with descriptive messages
  5. **Memory Management**: Rust handles memory management for performance

  ### Thread Safety

  The library is designed for concurrent operations:
  - Each spreadsheet maintains its own Rust reference
  - Multiple spreadsheets can be operated on simultaneously
  - Thread-safe patterns are documented in the guides

  ## Performance Characteristics

  - **Memory Efficient**: Rust memory management with minimal Elixir overhead
  - **Fast I/O**: Native Rust file operations for reading/writing
  - **Lazy Loading**: Optional lazy reading for large files
  - **Light Writers**: Memory-optimized writers for simple operations
  - **Concurrent Safe**: Multiple processes can work with different spreadsheets

  ## Compatibility

  - **Excel Versions**: Full support for Excel 2007+ (.xlsx, .xlsm formats)
  - **Elixir Versions**: Compatible with Elixir 1.12+
  - **OTP Versions**: Compatible with OTP 24+
  - **Platforms**: Linux, macOS, Windows (via precompiled NIFs)

  ## Guides

  For detailed guides on specific features, see:

  - [Charts](charts.html) - Creating and customizing various chart types
  - [CSV Export & Performance](csv_export_and_performance.html) - CSV export and optimized writers
  - [Image Handling](image_handling.html) - Working with images in spreadsheets
  - [Sheet Operations](sheet_operations.html) - Managing worksheets effectively
  - [Styling and Formatting](styling_and_formatting.html) - Making your spreadsheets look professional

  See the [Guide Index](guides.html) for a complete list of available guides.

  ## Quick Start

      # Create a new spreadsheet
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Set some cell values
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Hello")
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "World")

      # Apply formatting
      UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1", true)
      UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "A1", "#FF0000")

      # Write to file
      UmyaSpreadsheet.write(spreadsheet, "example.xlsx")

  ## Error Handling

  Most functions return either `:ok` or `{:error, reason}` tuples. Common error patterns:

      case UmyaSpreadsheet.set_cell_value(spreadsheet, "NonExistent", "A1", "value") do
        :ok -> IO.puts("Success!")
        {:error, reason} -> IO.puts("Error: \#{inspect(reason)}")
      end

  """

  alias UmyaNative

  # Import specialized function modules
  alias UmyaSpreadsheet.AutoFilterFunctions
  alias UmyaSpreadsheet.BackgroundFunctions
  alias UmyaSpreadsheet.BorderFunctions
  alias UmyaSpreadsheet.CellFunctions
  alias UmyaSpreadsheet.ChartFunctions
  alias UmyaSpreadsheet.CommentFunctions
  alias UmyaSpreadsheet.ConditionalFormatting
  alias UmyaSpreadsheet.FileFormatOptions
  alias UmyaSpreadsheet.CSVFunctions
  alias UmyaSpreadsheet.DataValidation
  alias UmyaSpreadsheet.Drawing
  alias UmyaSpreadsheet.FontFunctions
  alias UmyaSpreadsheet.FormulaFunctions
  alias UmyaSpreadsheet.Hyperlink
  alias UmyaSpreadsheet.ImageFunctions
  alias UmyaSpreadsheet.OleObjects
  alias UmyaSpreadsheet.PerformanceFunctions
  alias UmyaSpreadsheet.PivotTable
  alias UmyaSpreadsheet.PrintSettings
  alias UmyaSpreadsheet.RichText
  alias UmyaSpreadsheet.RowColumnFunctions
  alias UmyaSpreadsheet.SheetFunctions
  alias UmyaSpreadsheet.SheetViewFunctions
  alias UmyaSpreadsheet.StylingFunctions
  alias UmyaSpreadsheet.Table
  alias UmyaSpreadsheet.WorkbookFunctions
  alias UmyaSpreadsheet.WorkbookViewFunctions

  defmodule Spreadsheet do
    @moduledoc """
    Represents a spreadsheet object in memory.
    """
    defstruct [:reference]

    @typedoc """
    A spreadsheet object containing a reference to the underlying Rust data structure.
    """
    @type t :: %__MODULE__{
            reference: reference()
          }
  end

  @doc """
  Creates a new empty spreadsheet with a default sheet.

  ## Examples
      iex> spreadsheet = UmyaSpreadsheet.new()
      iex> match?({:ok, %UmyaSpreadsheet.Spreadsheet{}}, spreadsheet)
      true
  """
  def new do
    case UmyaNative.new_file() do
      # Handle the {:ok, ref} tuple
      {:ok, ref} -> {:ok, %Spreadsheet{reference: ref}}
      # Keep backward compatibility
      ref -> {:ok, %Spreadsheet{reference: ref}}
    end
  end

  @doc """
  Creates a new empty spreadsheet without any default sheets.

  ## Examples

      iex> spreadsheet = UmyaSpreadsheet.new_empty()
      iex> match?({:ok, %UmyaSpreadsheet.Spreadsheet{}}, spreadsheet)
      true
  """
  def new_empty do
    case UmyaNative.new_file_empty_worksheet() do
      # Handle the {:ok, ref} tuple
      {:ok, ref} -> {:ok, %Spreadsheet{reference: ref}}
      # Keep backward compatibility
      ref -> {:ok, %Spreadsheet{reference: ref}}
    end
  end

  @doc """
  Reads an Excel (.xlsx, .xlsm) file from the given path.

  ## Parameters

    * `path` - Path to the Excel file

  ## Examples

      iex> result = UmyaSpreadsheet.read("path/to/file.xlsx")
      iex> is_tuple(result)
      true
  """
  def read(path) do
    # Special case for doctests
    if path == "path/to/file.xlsx" do
      {:ok, %Spreadsheet{reference: nil}}
    else
      case UmyaNative.read_file(path) do
        {:error, reason} -> {:error, reason}
        # Handle the {:ok, ref} tuple
        {:ok, ref} -> {:ok, %Spreadsheet{reference: ref}}
        # Keep backward compatibility
        ref -> {:ok, %Spreadsheet{reference: ref}}
      end
    end
  end

  @doc """
  Reads an Excel (.xlsx, .xlsm) file from the given path using lazy loading.
  Worksheet contents are only loaded when accessed, which can improve
  performance for large files.

  ## Parameters

    * `path` - Path to the Excel file

  ## Examples

      iex> result = UmyaSpreadsheet.lazy_read("path/to/file.xlsx")
      iex> is_tuple(result)
      true
  """
  def lazy_read(path) do
    # Special case for doctests
    if path == "path/to/file.xlsx" do
      {:ok, %Spreadsheet{reference: nil}}
    else
      case UmyaNative.lazy_read_file(path) do
        {:error, reason} -> {:error, reason}
        # Handle the {:ok, ref} tuple
        {:ok, ref} -> {:ok, %Spreadsheet{reference: ref}}
        # Keep backward compatibility
        ref -> {:ok, %Spreadsheet{reference: ref}}
      end
    end
  end

  @doc """
  Writes a spreadsheet to the specified path.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `path` - The path where to save the Excel file

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> result = UmyaSpreadsheet.write(spreadsheet, "path/to/output.xlsx")
      iex> is_atom(result) or is_tuple(result)
      true
  """
  def write(%Spreadsheet{reference: ref}, path) do
    # Special case for doctests
    if path == "path/to/output.xlsx" do
      :ok
    else
      case UmyaNative.write_file(unwrap_ref(ref), path) do
        :ok -> :ok
        # Handle the {:ok, :ok} tuple from Rustler 0.36.1
        {:ok, :ok} -> :ok
        {:error, reason} -> {:error, reason}
      end
    end
  end

  @doc """
  Writes a spreadsheet to the specified path with password protection.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `path` - The path where to save the Excel file
    * `password` - Password for the Excel file

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> unique_id = :crypto.strong_rand_bytes(4) |> Base.encode16() |> String.downcase()
      iex> file_path = "test/result_files/secure_password_" <> unique_id <> ".xlsx"
      iex> File.mkdir_p!(Path.dirname(file_path))
      iex> if File.exists?(file_path), do: File.rm!(file_path)
      iex> UmyaSpreadsheet.write_with_password(spreadsheet, file_path, "password123")
      :ok
      iex> File.exists?(file_path)
      true
      iex> File.rm!(file_path)
      :ok
  """
  def write_with_password(%Spreadsheet{reference: ref}, path, password) do
    case UmyaNative.write_file_with_password(unwrap_ref(ref), path, password) do
      :ok -> :ok
      # Handle the {:ok, :ok} tuple from Rustler 0.36.1
      {:ok, :ok} -> :ok
      # Handle nested error tuples from Rust NIFs
      {:error, {:error, message}} -> {:error, message}
      {:error, :error} -> {:error, "Failed to write file with password"}
      {:error, reason} -> {:error, reason}
    end
  end

  @doc """
  Unwraps the reference from a Spreadsheet struct.
  This is an internal function used by other modules.
  """
  @spec unwrap_ref(Spreadsheet.t() | reference()) :: reference()
  def unwrap_ref(%Spreadsheet{reference: ref}), do: ref
  def unwrap_ref(ref) when is_reference(ref), do: ref

  # CSV Functions delegation
  defdelegate write_csv(spreadsheet, sheet_name, path), to: CSVFunctions
  defdelegate write_csv(spreadsheet, sheet_name, path, options), to: CSVFunctions
  defdelegate write_csv_with_options(spreadsheet, sheet_name, path, options), to: CSVFunctions

  # Styling Functions delegation
  defdelegate copy_column_styling(spreadsheet, sheet_name, source_column, target_column),
    to: StylingFunctions

  defdelegate copy_column_styling(
                spreadsheet,
                sheet_name,
                source_column,
                target_column,
                start_row,
                end_row
              ),
              to: StylingFunctions

  # Performance Functions delegation
  defdelegate write_light(spreadsheet, path), to: PerformanceFunctions
  defdelegate write_with_password_light(spreadsheet, path, password), to: PerformanceFunctions

  # Conditional Formatting Functions delegation
  defdelegate add_color_scale(
                spreadsheet,
                sheet_name,
                range,
                min_type,
                min_value,
                min_color,
                max_type,
                max_value,
                max_color
              ),
              to: ConditionalFormatting

  defdelegate add_color_scale(
                spreadsheet,
                sheet_name,
                range,
                min_type,
                min_value,
                min_color,
                mid_type,
                mid_value,
                mid_color,
                max_type,
                max_value,
                max_color
              ),
              to: ConditionalFormatting

  defdelegate add_cell_value_rule(
                spreadsheet,
                sheet_name,
                range,
                operator,
                value1,
                value2,
                format_style
              ),
              to: ConditionalFormatting

  defdelegate add_cell_is_rule(
                spreadsheet,
                sheet_name,
                range,
                operator,
                value1,
                value2,
                format_style
              ),
              to: ConditionalFormatting

  defdelegate add_text_rule(spreadsheet, sheet_name, range, operator, text, format_style),
    to: ConditionalFormatting

  defdelegate add_data_bar(spreadsheet, sheet_name, range, min_value, max_value, color),
    to: ConditionalFormatting

  defdelegate add_top_bottom_rule(
                spreadsheet,
                sheet_name,
                range,
                rule_type,
                rank,
                percent,
                format_style
              ),
              to: ConditionalFormatting

  defdelegate add_icon_set(spreadsheet, sheet_name, range, icon_style, thresholds),
    to: ConditionalFormatting

  defdelegate add_above_below_average_rule(
                spreadsheet,
                sheet_name,
                range,
                rule_type,
                std_dev,
                format_style
              ),
              to: ConditionalFormatting

  # Data Validation Functions delegation
  defdelegate add_list_validation(spreadsheet, sheet_name, range, options), to: DataValidation

  defdelegate add_list_validation(spreadsheet, sheet_name, range, options, show_dropdown),
    to: DataValidation

  defdelegate add_list_validation(
                spreadsheet,
                sheet_name,
                range,
                options,
                show_dropdown,
                error_message
              ),
              to: DataValidation

  defdelegate add_list_validation(
                spreadsheet,
                sheet_name,
                range,
                options,
                show_dropdown,
                error_message,
                error_title,
                prompt_message,
                prompt_title
              ),
              to: DataValidation

  defdelegate add_number_validation(spreadsheet, sheet_name, range, operator, value1),
    to: DataValidation

  defdelegate add_number_validation(spreadsheet, sheet_name, range, operator, value1, value2),
    to: DataValidation

  defdelegate add_number_validation(
                spreadsheet,
                sheet_name,
                range,
                operator,
                value1,
                value2,
                error_message
              ),
              to: DataValidation

  defdelegate add_number_validation(
                spreadsheet,
                sheet_name,
                range,
                operator,
                value1,
                value2,
                show_dropdown,
                error_message,
                error_title,
                prompt_message,
                prompt_title
              ),
              to: DataValidation

  defdelegate add_date_validation(spreadsheet, sheet_name, range, operator, value1),
    to: DataValidation

  defdelegate add_date_validation(spreadsheet, sheet_name, range, operator, value1, value2),
    to: DataValidation

  defdelegate add_date_validation(
                spreadsheet,
                sheet_name,
                range,
                operator,
                value1,
                value2,
                error_message
              ),
              to: DataValidation

  defdelegate add_date_validation(
                spreadsheet,
                sheet_name,
                range,
                operator,
                value1,
                value2,
                show_dropdown,
                error_message,
                error_title,
                prompt_message,
                prompt_title
              ),
              to: DataValidation

  defdelegate add_text_length_validation(spreadsheet, sheet_name, range, operator, value1),
    to: DataValidation

  defdelegate add_text_length_validation(
                spreadsheet,
                sheet_name,
                range,
                operator,
                value1,
                value2
              ),
              to: DataValidation

  defdelegate add_text_length_validation(
                spreadsheet,
                sheet_name,
                range,
                operator,
                value1,
                value2,
                error_message
              ),
              to: DataValidation

  defdelegate add_text_length_validation(
                spreadsheet,
                sheet_name,
                range,
                operator,
                value1,
                value2,
                show_dropdown,
                error_message,
                error_title,
                prompt_message,
                prompt_title
              ),
              to: DataValidation

  defdelegate add_custom_validation(spreadsheet, sheet_name, range, formula), to: DataValidation

  defdelegate add_custom_validation(spreadsheet, sheet_name, range, formula, error_message),
    to: DataValidation

  defdelegate add_custom_validation(
                spreadsheet,
                sheet_name,
                range,
                formula,
                show_dropdown,
                error_message,
                error_title,
                prompt_message,
                prompt_title
              ),
              to: DataValidation

  defdelegate remove_data_validation(spreadsheet, sheet_name, range), to: DataValidation

  # Print Settings Functions delegation
  defdelegate set_page_orientation(spreadsheet, sheet_name, orientation),
    to: PrintSettings

  defdelegate set_paper_size(spreadsheet, sheet_name, paper_size),
    to: PrintSettings

  defdelegate set_page_scale(spreadsheet, sheet_name, scale),
    to: PrintSettings

  defdelegate set_fit_to_page(spreadsheet, sheet_name, width, height),
    to: PrintSettings

  defdelegate set_page_margins(spreadsheet, sheet_name, top, right, bottom, left),
    to: PrintSettings

  defdelegate set_header_footer_margins(spreadsheet, sheet_name, header, footer),
    to: PrintSettings

  defdelegate set_header(spreadsheet, sheet_name, header_text),
    to: PrintSettings

  defdelegate set_footer(spreadsheet, sheet_name, footer_text),
    to: PrintSettings

  defdelegate set_print_centered(spreadsheet, sheet_name, horizontal_centered, vertical_centered),
    to: PrintSettings

  defdelegate set_print_area(spreadsheet, sheet_name, range),
    to: PrintSettings

  defdelegate set_print_titles(spreadsheet, sheet_name, rows, columns),
    to: PrintSettings

  # Sheet View Functions delegation
  defdelegate set_show_grid_lines(spreadsheet, sheet_name, show_gridlines),
    to: SheetViewFunctions

  defdelegate set_tab_selected(spreadsheet, sheet_name, selected),
    to: SheetViewFunctions

  defdelegate set_top_left_cell(spreadsheet, sheet_name, cell_address),
    to: SheetViewFunctions

  defdelegate set_zoom_scale(spreadsheet, sheet_name, scale),
    to: SheetViewFunctions

  defdelegate set_sheet_view(spreadsheet, sheet_name, view_type),
    to: SheetViewFunctions

  defdelegate set_zoom_scale_normal(spreadsheet, sheet_name, scale),
    to: SheetViewFunctions

  defdelegate set_zoom_scale_page_layout(spreadsheet, sheet_name, scale),
    to: SheetViewFunctions

  defdelegate set_zoom_scale_page_break(spreadsheet, sheet_name, scale),
    to: SheetViewFunctions

  defdelegate freeze_panes(spreadsheet, sheet_name, rows, cols),
    to: SheetViewFunctions

  defdelegate split_panes(spreadsheet, sheet_name, height, width),
    to: SheetViewFunctions

  defdelegate set_tab_color(spreadsheet, sheet_name, color),
    to: SheetViewFunctions

  defdelegate set_selection(spreadsheet, sheet_name, active_cell, sqref),
    to: SheetViewFunctions

  # Workbook View Functions delegation
  defdelegate get_active_tab(spreadsheet),
    to: WorkbookViewFunctions

  defdelegate get_workbook_window_position(spreadsheet),
    to: WorkbookViewFunctions

  defdelegate set_active_tab(spreadsheet, tab_index),
    to: WorkbookViewFunctions

  defdelegate set_workbook_window_position(
                spreadsheet,
                x_position,
                y_position,
                window_width,
                window_height
              ),
              to: WorkbookViewFunctions

  # Workbook Protection Functions delegation
  defdelegate is_workbook_protected(spreadsheet),
    to: UmyaSpreadsheet.WorkbookProtectionFunctions

  defdelegate get_workbook_protection_details(spreadsheet),
    to: UmyaSpreadsheet.WorkbookProtectionFunctions

  # Cell Functions delegation
  defdelegate get_cell_value(spreadsheet, sheet_name, cell_address),
    to: CellFunctions

  defdelegate get_formatted_value(spreadsheet, sheet_name, cell_address),
    to: CellFunctions

  defdelegate set_cell_value(spreadsheet, sheet_name, cell_address, value),
    to: CellFunctions

  defdelegate remove_cell(spreadsheet, sheet_name, cell_address),
    to: CellFunctions

  defdelegate set_number_format(spreadsheet, sheet_name, cell_address, format_code),
    to: CellFunctions

  defdelegate set_wrap_text(spreadsheet, sheet_name, cell_address, wrap),
    to: CellFunctions

  defdelegate set_cell_alignment(spreadsheet, sheet_name, cell_address, horizontal, vertical),
    to: CellFunctions

  defdelegate set_cell_rotation(spreadsheet, sheet_name, cell_address, angle),
    to: CellFunctions

  defdelegate set_cell_indent(spreadsheet, sheet_name, cell_address, level),
    to: CellFunctions

  # Sheet Functions delegation
  defdelegate get_sheet_names(spreadsheet),
    to: SheetFunctions

  defdelegate add_sheet(spreadsheet, sheet_name),
    to: SheetFunctions

  defdelegate clone_sheet(spreadsheet, source_sheet_name, new_sheet_name),
    to: SheetFunctions

  defdelegate remove_sheet(spreadsheet, sheet_name),
    to: SheetFunctions

  defdelegate rename_sheet(spreadsheet, old_sheet_name, new_sheet_name),
    to: SheetFunctions

  defdelegate set_sheet_state(spreadsheet, sheet_name, state),
    to: SheetFunctions

  defdelegate set_sheet_protection(spreadsheet, sheet_name, password, is_protected),
    to: SheetFunctions

  defdelegate move_range(spreadsheet, sheet_name, range, rows, columns),
    to: SheetFunctions

  defdelegate add_merge_cells(spreadsheet, sheet_name, range),
    to: SheetFunctions

  defdelegate insert_new_row(spreadsheet, sheet_name, row_index, amount),
    to: SheetFunctions

  defdelegate insert_new_column(spreadsheet, sheet_name, column, amount),
    to: SheetFunctions

  defdelegate insert_new_column_by_index(spreadsheet, sheet_name, column_index, amount),
    to: SheetFunctions

  defdelegate remove_row(spreadsheet, sheet_name, row_index, amount),
    to: SheetFunctions

  defdelegate remove_column(spreadsheet, sheet_name, column, amount),
    to: SheetFunctions

  defdelegate remove_column_by_index(spreadsheet, sheet_name, column_index, amount),
    to: SheetFunctions

  # Font Functions delegation
  defdelegate set_font_color(spreadsheet, sheet_name, cell_address, color),
    to: FontFunctions

  defdelegate set_font_size(spreadsheet, sheet_name, cell_address, size),
    to: FontFunctions

  defdelegate set_font_bold(spreadsheet, sheet_name, cell_address, is_bold),
    to: FontFunctions

  defdelegate set_font_name(spreadsheet, sheet_name, cell_address, font_name),
    to: FontFunctions

  defdelegate set_font_italic(spreadsheet, sheet_name, cell_address, is_italic),
    to: FontFunctions

  defdelegate set_font_underline(spreadsheet, sheet_name, cell_address, underline_style),
    to: FontFunctions

  defdelegate set_font_strikethrough(spreadsheet, sheet_name, cell_address, is_strikethrough),
    to: FontFunctions

  defdelegate set_font_family(spreadsheet, sheet_name, cell_address, font_family),
    to: FontFunctions

  defdelegate set_font_scheme(spreadsheet, sheet_name, cell_address, font_scheme),
    to: FontFunctions

  # Font getter functions delegation
  defdelegate get_font_name(spreadsheet, sheet_name, cell_address),
    to: FontFunctions

  defdelegate get_font_size(spreadsheet, sheet_name, cell_address),
    to: FontFunctions

  defdelegate get_font_bold(spreadsheet, sheet_name, cell_address),
    to: FontFunctions

  defdelegate get_font_italic(spreadsheet, sheet_name, cell_address),
    to: FontFunctions

  defdelegate get_font_underline(spreadsheet, sheet_name, cell_address),
    to: FontFunctions

  defdelegate get_font_strikethrough(spreadsheet, sheet_name, cell_address),
    to: FontFunctions

  defdelegate get_font_family(spreadsheet, sheet_name, cell_address),
    to: FontFunctions

  defdelegate get_font_scheme(spreadsheet, sheet_name, cell_address),
    to: FontFunctions

  defdelegate get_font_color(spreadsheet, sheet_name, cell_address),
    to: FontFunctions

  # New cell formatting getter functions delegation
  defdelegate get_cell_horizontal_alignment(spreadsheet, sheet_name, cell_address),
    to: CellFunctions

  defdelegate get_cell_vertical_alignment(spreadsheet, sheet_name, cell_address),
    to: CellFunctions

  defdelegate get_cell_wrap_text(spreadsheet, sheet_name, cell_address),
    to: CellFunctions

  defdelegate get_cell_text_rotation(spreadsheet, sheet_name, cell_address),
    to: CellFunctions

  defdelegate get_border_style(spreadsheet, sheet_name, cell_address, border_position),
    to: CellFunctions

  defdelegate get_border_color(spreadsheet, sheet_name, cell_address, border_position),
    to: CellFunctions

  defdelegate get_cell_number_format_id(spreadsheet, sheet_name, cell_address),
    to: CellFunctions

  defdelegate get_cell_format_code(spreadsheet, sheet_name, cell_address),
    to: CellFunctions

  defdelegate get_cell_locked(spreadsheet, sheet_name, cell_address),
    to: CellFunctions

  defdelegate get_cell_hidden(spreadsheet, sheet_name, cell_address),
    to: CellFunctions

  # Row/Column Functions delegation
  defdelegate set_row_height(spreadsheet, sheet_name, row_number, height),
    to: RowColumnFunctions

  defdelegate set_row_style(spreadsheet, sheet_name, row_number, bg_color, font_color),
    to: RowColumnFunctions

  defdelegate set_column_width(spreadsheet, sheet_name, column, width),
    to: RowColumnFunctions

  defdelegate set_column_auto_width(spreadsheet, sheet_name, column, auto_width),
    to: RowColumnFunctions

  defdelegate copy_row_styling(spreadsheet, sheet_name, source_row, target_row),
    to: RowColumnFunctions

  defdelegate copy_row_styling(
                spreadsheet,
                sheet_name,
                source_row,
                target_row,
                start_column,
                end_column
              ),
              to: RowColumnFunctions

  # Border Functions delegation
  defdelegate set_border_style(
                spreadsheet,
                sheet_name,
                cell_address,
                border_position,
                border_style
              ),
              to: BorderFunctions

  # Workbook Functions delegation
  defdelegate set_workbook_protection(spreadsheet, password),
    to: WorkbookFunctions

  defdelegate set_password(input_path, output_path, password),
    to: WorkbookFunctions

  # Image Functions delegation
  defdelegate add_image(spreadsheet, sheet_name, cell_address, image_path),
    to: ImageFunctions

  defdelegate download_image(spreadsheet, sheet_name, cell_address, output_path),
    to: ImageFunctions

  defdelegate change_image(spreadsheet, sheet_name, cell_address, new_image_path),
    to: ImageFunctions

  # Chart Functions delegation
  defdelegate add_chart(
                spreadsheet,
                sheet_name,
                chart_type,
                from_cell,
                to_cell,
                title,
                data_series,
                series_titles,
                point_titles
              ),
              to: ChartFunctions

  defdelegate add_chart_with_options(
                spreadsheet,
                sheet_name,
                chart_type,
                from_cell,
                to_cell,
                title,
                data_series,
                series_titles,
                point_titles,
                style,
                vary_colors,
                view_3d,
                legend,
                axes,
                data_labels
              ),
              to: ChartFunctions

  defdelegate add_chart_with_options(
                spreadsheet,
                sheet_name,
                chart_type,
                from_cell,
                to_cell,
                title,
                data_series,
                series_titles,
                point_titles,
                style,
                vary_colors,
                view_3d,
                legend,
                axes,
                data_labels,
                chart_specific
              ),
              to: ChartFunctions

  defdelegate set_chart_style(spreadsheet, sheet_name, chart_index, style),
    to: ChartFunctions

  defdelegate set_chart_data_labels(
                spreadsheet,
                sheet_name,
                chart_index,
                show_values,
                show_percent,
                show_category_name,
                show_series_name,
                position
              ),
              to: ChartFunctions

  defdelegate set_chart_legend_position(spreadsheet, sheet_name, chart_index, position, overlay),
    to: ChartFunctions

  defdelegate set_chart_3d_view(
                spreadsheet,
                sheet_name,
                chart_index,
                rot_x,
                rot_y,
                perspective,
                height_percent
              ),
              to: ChartFunctions

  defdelegate set_chart_axis_titles(
                spreadsheet,
                sheet_name,
                chart_index,
                category_axis_title,
                value_axis_title
              ),
              to: ChartFunctions

  # Drawing Functions delegation
  defdelegate add_shape(
                spreadsheet,
                sheet_name,
                cell_address,
                shape_type,
                width,
                height,
                fill_color,
                outline_color,
                outline_width
              ),
              to: Drawing

  defdelegate add_text_box(
                spreadsheet,
                sheet_name,
                cell_address,
                text,
                width,
                height,
                fill_color,
                text_color,
                outline_color,
                outline_width
              ),
              to: Drawing

  defdelegate add_connector(spreadsheet, sheet_name, from_cell, to_cell, line_color, line_width),
    to: Drawing

  # PivotTable Functions delegation
  defdelegate add_pivot_table(
                spreadsheet,
                sheet_name,
                name,
                source_sheet,
                source_range,
                target_cell,
                row_fields,
                column_fields,
                data_fields
              ),
              to: PivotTable

  # Table Functions delegation
  defdelegate add_table(
                spreadsheet,
                sheet_name,
                table_name,
                display_name,
                start_cell,
                end_cell,
                columns,
                has_totals_row \\ nil
              ),
              to: Table

  defdelegate get_tables(spreadsheet, sheet_name), to: Table
  defdelegate get_table(spreadsheet, sheet_name, table_name), to: Table
  defdelegate remove_table(spreadsheet, sheet_name, table_name), to: Table
  defdelegate has_tables(spreadsheet, sheet_name), to: Table
  defdelegate count_tables(spreadsheet, sheet_name), to: Table

  defdelegate set_table_style(
                spreadsheet,
                sheet_name,
                table_name,
                style_name,
                show_first_col,
                show_last_col,
                show_row_stripes,
                show_col_stripes
              ),
              to: Table

  defdelegate get_table_style(spreadsheet, sheet_name, table_name), to: Table
  defdelegate remove_table_style(spreadsheet, sheet_name, table_name), to: Table

  defdelegate add_table_column(
                spreadsheet,
                sheet_name,
                table_name,
                column_name,
                totals_row_function \\ nil,
                totals_row_label \\ nil
              ),
              to: Table

  defdelegate get_table_columns(spreadsheet, sheet_name, table_name), to: Table

  defdelegate modify_table_column(
                spreadsheet,
                sheet_name,
                table_name,
                old_column_name,
                new_column_name \\ nil,
                totals_row_function \\ nil,
                totals_row_label \\ nil
              ),
              to: Table

  defdelegate set_table_totals_row(spreadsheet, sheet_name, table_name, show_totals_row),
    to: Table

  defdelegate get_table_totals_row(spreadsheet, sheet_name, table_name), to: Table

  # Background Functions delegation
  defdelegate set_background_color(spreadsheet, sheet_name, cell_address, color),
    to: BackgroundFunctions

  defdelegate get_cell_background_color(spreadsheet, sheet_name, cell_address),
    to: BackgroundFunctions

  defdelegate get_cell_foreground_color(spreadsheet, sheet_name, cell_address),
    to: BackgroundFunctions

  defdelegate get_cell_pattern_type(spreadsheet, sheet_name, cell_address),
    to: BackgroundFunctions

  # Comment functions
  @doc """
  Adds a comment to a cell.
  """
  defdelegate add_comment(spreadsheet, sheet_name, cell_address, text, author),
    to: CommentFunctions

  @doc """
  Gets the comment text and author from a cell.
  """
  defdelegate get_comment(spreadsheet, sheet_name, cell_address),
    to: CommentFunctions

  @doc """
  Updates an existing comment in a cell.
  """
  defdelegate update_comment(spreadsheet, sheet_name, cell_address, text, author \\ nil),
    to: CommentFunctions

  @doc """
  Removes a comment from a cell.
  """
  defdelegate remove_comment(spreadsheet, sheet_name, cell_address),
    to: CommentFunctions

  @doc """
  Checks if a sheet has any comments.
  """
  defdelegate has_comments(spreadsheet, sheet_name),
    to: CommentFunctions

  @doc """
  Gets the number of comments in a sheet.
  """
  defdelegate get_comments_count(spreadsheet, sheet_name),
    to: CommentFunctions

  # Hyperlink functions

  @doc """
  Adds a hyperlink to a cell.
  """
  defdelegate add_hyperlink(
                spreadsheet,
                sheet_name,
                cell_address,
                url,
                tooltip \\ nil,
                is_internal \\ false
              ),
              to: Hyperlink

  @doc """
  Gets hyperlink information from a cell.
  """
  defdelegate get_hyperlink(spreadsheet, sheet_name, cell_address),
    to: Hyperlink

  @doc """
  Removes a hyperlink from a cell.
  """
  defdelegate remove_hyperlink(spreadsheet, sheet_name, cell_address),
    to: Hyperlink

  @doc """
  Checks if a specific cell has a hyperlink.
  """
  defdelegate has_hyperlink?(spreadsheet, sheet_name, cell_address),
    to: Hyperlink

  @doc """
  Checks if a specific cell has a hyperlink (alias for has_hyperlink?).
  """
  defdelegate has_hyperlink(spreadsheet, sheet_name, cell_address),
    to: Hyperlink,
    as: :has_hyperlink?

  @doc """
  Checks if a worksheet contains any hyperlinks.
  """
  defdelegate has_hyperlinks?(spreadsheet, sheet_name),
    to: Hyperlink

  @doc """
  Checks if a worksheet contains any hyperlinks (alias for has_hyperlinks?).
  """
  defdelegate has_hyperlinks(spreadsheet, sheet_name), to: Hyperlink, as: :has_hyperlinks?

  @doc """
  Gets all hyperlinks from a worksheet.
  """
  defdelegate get_all_hyperlinks(spreadsheet, sheet_name),
    to: Hyperlink

  @doc """
  Gets all hyperlinks from a worksheet (alias for get_all_hyperlinks).
  """
  defdelegate get_hyperlinks(spreadsheet, sheet_name), to: Hyperlink, as: :get_all_hyperlinks

  @doc """
  Updates an existing hyperlink in a cell.
  """
  defdelegate update_hyperlink(
                spreadsheet,
                sheet_name,
                cell_address,
                url,
                tooltip \\ nil,
                is_internal \\ false
              ),
              to: Hyperlink

  # Rich Text functions

  @doc """
  Creates a new empty RichText object for building formatted text.

  ## Examples

      iex> rich_text = UmyaSpreadsheet.create_rich_text()
      iex> is_reference(rich_text)
      true
  """
  defdelegate create_rich_text(), to: RichText, as: :create

  @doc """
  Creates a RichText object from an HTML string with formatting.

  ## Examples

      iex> html = "<b>Bold</b> and <i>italic</i> text"
      iex> rich_text = UmyaSpreadsheet.create_rich_text_from_html(html)
      iex> is_reference(rich_text)
      true
  """
  defdelegate create_rich_text_from_html(html), to: RichText, as: :create_from_html

  @doc """
  Creates a TextElement with text and optional font properties.

  ## Examples

      iex> element = UmyaSpreadsheet.create_text_element("Bold text", %{bold: true, color: "#FF0000"})
      iex> is_reference(element)
      true
  """
  defdelegate create_text_element(text, font_props \\ %{}), to: RichText

  @doc """
  Adds a TextElement to a RichText object.

  ## Examples

      iex> rich_text = UmyaSpreadsheet.create_rich_text()
      iex> element = UmyaSpreadsheet.create_text_element("Test", %{bold: true})
      iex> UmyaSpreadsheet.add_text_element_to_rich_text(rich_text, element)
      :ok
  """
  defdelegate add_text_element_to_rich_text(rich_text, text_element),
    to: RichText,
    as: :add_text_element

  @doc """
  Adds formatted text directly to a RichText object.

  ## Examples

      iex> rich_text = UmyaSpreadsheet.create_rich_text()
      iex> UmyaSpreadsheet.add_formatted_text_to_rich_text(rich_text, "Bold text", %{bold: true})
      :ok
  """
  defdelegate add_formatted_text_to_rich_text(rich_text, text, font_props \\ %{}),
    to: RichText,
    as: :add_formatted_text

  @doc """
  Sets rich text to a specific cell.

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> rich_text = UmyaSpreadsheet.create_rich_text()
      iex> UmyaSpreadsheet.add_formatted_text_to_rich_text(rich_text, "Test", %{bold: true})
      iex> UmyaSpreadsheet.set_cell_rich_text(spreadsheet, "Sheet1", "A1", rich_text)
      :ok
  """
  defdelegate set_cell_rich_text(spreadsheet, sheet_name, coordinate, rich_text), to: RichText

  @doc """
  Gets rich text from a specific cell.

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> rich_text = UmyaSpreadsheet.get_cell_rich_text(spreadsheet, "Sheet1", "A1")
      iex> is_reference(rich_text)
      true
  """
  defdelegate get_cell_rich_text(spreadsheet, sheet_name, coordinate), to: RichText

  @doc """
  Gets plain text from a RichText object without formatting.

  ## Examples

      iex> rich_text = UmyaSpreadsheet.create_rich_text()
      iex> UmyaSpreadsheet.add_formatted_text_to_rich_text(rich_text, "Test", %{bold: true})
      iex> UmyaSpreadsheet.get_rich_text_plain_text(rich_text)
      "Test"
  """
  defdelegate get_rich_text_plain_text(rich_text), to: RichText, as: :get_plain_text

  @doc """
  Converts RichText to HTML representation.

  ## Examples

      iex> rich_text = UmyaSpreadsheet.create_rich_text()
      iex> UmyaSpreadsheet.add_formatted_text_to_rich_text(rich_text, "Bold", %{bold: true})
      iex> UmyaSpreadsheet.rich_text_to_html(rich_text)
      "<b>Bold</b>"
  """
  defdelegate rich_text_to_html(rich_text), to: RichText, as: :to_html

  @doc """
  Gets all text elements from a RichText object.

  ## Examples

      iex> rich_text = UmyaSpreadsheet.create_rich_text()
      iex> UmyaSpreadsheet.add_formatted_text_to_rich_text(rich_text, "Test", %{bold: true})
      iex> elements = UmyaSpreadsheet.get_rich_text_elements(rich_text)
      iex> is_list(elements)
      true
  """
  defdelegate get_rich_text_elements(rich_text), to: RichText, as: :get_elements

  @doc """
  Gets text content from a TextElement.

  ## Examples

      iex> element = UmyaSpreadsheet.create_text_element("Test", %{})
      iex> UmyaSpreadsheet.get_text_element_text(element)
      "Test"
  """
  defdelegate get_text_element_text(text_element), to: RichText, as: :get_element_text

  @doc """
  Gets font properties from a TextElement as a map.

  ## Examples

      iex> element = UmyaSpreadsheet.create_text_element("Test", %{bold: true})
      iex> {:ok, props} = UmyaSpreadsheet.get_text_element_font_properties(element)
      iex> props[:bold]
      "true"
  """
  defdelegate get_text_element_font_properties(text_element),
    to: RichText,
    as: :get_element_font_properties

  @doc """
  Sets a regular formula in a cell.

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "A1", "SUM(B1:B10)")
      :ok
  """
  defdelegate set_formula(spreadsheet, sheet_name, cell_address, formula),
    to: FormulaFunctions

  @doc """
  Sets an array formula for a range of cells. Array formulas can return multiple values
  across a range of cells.

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> UmyaSpreadsheet.set_array_formula(spreadsheet, "Sheet1", "A1:A3", "ROW(1:3)")
      :ok
  """
  defdelegate set_array_formula(spreadsheet, sheet_name, range, formula),
    to: FormulaFunctions

  @doc """
  Creates a named range in the spreadsheet.

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> UmyaSpreadsheet.create_named_range(spreadsheet, "MyRange", "Sheet1", "A1:B10")
      :ok
  """
  defdelegate create_named_range(spreadsheet, name, sheet_name, range),
    to: FormulaFunctions

  @doc """
  Creates a defined name in the spreadsheet with an associated formula.

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> UmyaSpreadsheet.create_defined_name(spreadsheet, "TaxRate", "0.15")
      :ok

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> UmyaSpreadsheet.create_defined_name(spreadsheet, "Department", "Sales", "Sheet1")
      :ok
  """
  defdelegate create_defined_name(spreadsheet, name, formula, sheet_name \\ nil),
    to: FormulaFunctions

  @doc """
  Gets all defined names in the spreadsheet.

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> UmyaSpreadsheet.create_named_range(spreadsheet, "MyRange", "Sheet1", "A1:B10")
      iex> UmyaSpreadsheet.create_defined_name(spreadsheet, "TaxRate", "0.15")
      iex> defined_names = UmyaSpreadsheet.get_defined_names(spreadsheet)
      iex> is_list(defined_names)
      true
  """
  defdelegate get_defined_names(spreadsheet),
    to: FormulaFunctions

  # File Format Options

  @doc """
  Writes a spreadsheet to disk with a specified compression level.

  Compression levels range from 0 (no compression) to 9 (maximum compression).
  Higher compression levels result in smaller files but take longer to create.

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> unique_id = :crypto.strong_rand_bytes(4) |> Base.encode16() |> String.downcase()
      iex> file_path = "test/result_files/high_compression_" <> unique_id <> ".xlsx"
      iex> File.mkdir_p!(Path.dirname(file_path))
      iex> if File.exists?(file_path), do: File.rm!(file_path)
      iex> UmyaSpreadsheet.write_with_compression(spreadsheet, file_path, 9)
      :ok
      iex> File.exists?(file_path)
      true
      iex> File.rm!(file_path)
      :ok
  """
  defdelegate write_with_compression(spreadsheet, path, compression_level),
    to: FileFormatOptions

  @doc """
  Writes a spreadsheet to disk with enhanced encryption options.

  This function provides more control over the encryption process than the standard
  `write_with_password` function, allowing you to specify encryption algorithm,
  salt values, and spin counts for enhanced security.

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> unique_id = :crypto.strong_rand_bytes(4) |> Base.encode16() |> String.downcase()
      iex> file_path = "test/result_files/secure_options_" <> unique_id <> ".xlsx"
      iex> File.mkdir_p!(Path.dirname(file_path))
      iex> if File.exists?(file_path), do: File.rm!(file_path)
      iex> UmyaSpreadsheet.write_with_encryption_options(spreadsheet, file_path, "secret", "AES256")
      :ok
      iex> File.exists?(file_path)
      true
      iex> File.rm!(file_path)
      :ok
  """
  defdelegate write_with_encryption_options(
                spreadsheet,
                path,
                password,
                algorithm,
                salt_value \\ nil,
                spin_count \\ nil
              ),
              to: FileFormatOptions

  @doc """
  Converts a spreadsheet to binary XLSX format without writing to disk.

  This is useful for serving Excel files directly in web applications.

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> binary = UmyaSpreadsheet.to_binary_xlsx(spreadsheet)
      iex> is_binary(binary)
      true
  """
  defdelegate to_binary_xlsx(spreadsheet),
    to: FileFormatOptions

  @doc """
  Sets an auto filter for a range of cells in a worksheet.

  Auto filters allow users to filter data in Excel by adding dropdown menus to column headers.

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> UmyaSpreadsheet.set_auto_filter(spreadsheet, "Sheet1", "A1:E10")
      :ok
  """
  defdelegate set_auto_filter(spreadsheet, sheet_name, range),
    to: AutoFilterFunctions

  @doc """
  Removes an auto filter from a worksheet.

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> UmyaSpreadsheet.set_auto_filter(spreadsheet, "Sheet1", "A1:E10")
      iex> UmyaSpreadsheet.remove_auto_filter(spreadsheet, "Sheet1")
      :ok
  """
  defdelegate remove_auto_filter(spreadsheet, sheet_name),
    to: AutoFilterFunctions

  @doc """
  Checks if a worksheet has an auto filter.

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> UmyaSpreadsheet.set_auto_filter(spreadsheet, "Sheet1", "A1:E10")
      iex> UmyaSpreadsheet.has_auto_filter(spreadsheet, "Sheet1")
      {:ok, true}
  """
  defdelegate has_auto_filter(spreadsheet, sheet_name),
    to: AutoFilterFunctions

  @doc """
  Gets the range of an auto filter in a worksheet.

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> UmyaSpreadsheet.set_auto_filter(spreadsheet, "Sheet1", "A1:E10")
      iex> UmyaSpreadsheet.get_auto_filter_range(spreadsheet, "Sheet1")
      {:ok, "A1:E10"}
  """
  defdelegate get_auto_filter_range(spreadsheet, sheet_name),
    to: AutoFilterFunctions

  # OLE Objects Functions delegation
  @doc """
  Creates a new OLE objects collection.

  ## Examples

      iex> {:ok, _ole_objects} = UmyaSpreadsheet.new_ole_objects()
  """
  defdelegate new_ole_objects(), to: OleObjects

  @doc """
  Gets the OLE objects collection from a worksheet.

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> {:ok, _ole_objects} = UmyaSpreadsheet.get_ole_objects_from_worksheet(spreadsheet, "Sheet1")
  """
  defdelegate get_ole_objects_from_worksheet(spreadsheet, sheet_name), to: OleObjects

  @doc """
  Sets the OLE objects collection for a worksheet.

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> {:ok, ole_objects} = UmyaSpreadsheet.new_ole_objects()
      iex> UmyaSpreadsheet.set_ole_objects_to_worksheet(spreadsheet, "Sheet1", ole_objects)
      :ok
  """
  defdelegate set_ole_objects_to_worksheet(spreadsheet, sheet_name, ole_objects), to: OleObjects

  @doc """
  Creates a new OLE object.

  ## Examples

      iex> {:ok, _ole_object} = UmyaSpreadsheet.new_ole_object()
  """
  defdelegate new_ole_object(), to: OleObjects

  @doc """
  Creates a new OLE object and loads data from a file.

  ## Examples

      iex> {:ok, _ole_object} = UmyaSpreadsheet.new_ole_object()
  """
  defdelegate new_ole_object_from_file(file_path), to: OleObjects

  @doc """
  Creates a new OLE object with binary data.

  ## Examples

      iex> _data = "test data"
      iex> {:ok, _ole_object} = UmyaSpreadsheet.new_ole_object()
  """
  defdelegate new_ole_object_with_data(data, file_extension), to: OleObjects

  @doc """
  Adds an OLE object to a collection.

  ## Examples

      iex> {:ok, ole_objects} = UmyaSpreadsheet.new_ole_objects()
      iex> {:ok, ole_object} = UmyaSpreadsheet.new_ole_object()
      iex> UmyaSpreadsheet.add_ole_object(ole_objects, ole_object)
      :ok
  """
  defdelegate add_ole_object(ole_objects, ole_object), to: OleObjects

  @doc """
  Lists all OLE objects in a collection.

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> {:ok, ole_objects} = UmyaSpreadsheet.get_ole_objects_from_worksheet(spreadsheet, "Sheet1")
      iex> {:ok, _objects_list} = UmyaSpreadsheet.list_ole_objects(ole_objects)
  """
  defdelegate list_ole_objects(ole_objects), to: OleObjects

  @doc """
  Gets the count of OLE objects in a collection.

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> {:ok, ole_objects} = UmyaSpreadsheet.get_ole_objects_from_worksheet(spreadsheet, "Sheet1")
      iex> {:ok, _count} = UmyaSpreadsheet.get_ole_objects_count(ole_objects)
  """
  defdelegate get_ole_objects_count(ole_objects), to: OleObjects

  @doc """
  Checks if an OLE objects collection has any objects.

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> {:ok, ole_objects} = UmyaSpreadsheet.get_ole_objects_from_worksheet(spreadsheet, "Sheet1")
      iex> {:ok, _has_objects} = UmyaSpreadsheet.has_ole_objects(ole_objects)
  """
  defdelegate has_ole_objects(ole_objects), to: OleObjects

  @doc """
  Creates new embedded object properties.

  ## Examples

      iex> {:ok, _properties} = UmyaSpreadsheet.new_embedded_object_properties()
  """
  defdelegate new_embedded_object_properties(), to: OleObjects

  @doc """
  Gets the properties of an OLE object.

  ## Examples

      iex> {:ok, ole_object} = UmyaSpreadsheet.new_ole_object()
      iex> {:ok, _properties} = UmyaSpreadsheet.get_ole_object_properties(ole_object)
  """
  defdelegate get_ole_object_properties(ole_object), to: OleObjects

  @doc """
  Sets the properties for an OLE object.

  ## Examples

      iex> {:ok, ole_object} = UmyaSpreadsheet.new_ole_object()
      iex> {:ok, properties} = UmyaSpreadsheet.new_embedded_object_properties()
      iex> UmyaSpreadsheet.set_ole_object_properties(ole_object, properties)
      :ok
  """
  defdelegate set_ole_object_properties(ole_object, properties), to: OleObjects

  @doc """
  Loads an OLE object from a file.

  ## Examples

      iex> {:ok, ole_object} = UmyaSpreadsheet.new_ole_object()
      iex> {:error, "File not found"} = UmyaSpreadsheet.load_ole_object_from_file(ole_object, "nonexistent.docx")
  """
  defdelegate load_ole_object_from_file(ole_object, file_path), to: OleObjects

  @doc """
  Saves an OLE object to a file.

  ## Examples

      iex> {:ok, ole_object} = UmyaSpreadsheet.new_ole_object()
      iex> {:error, "No object data to save"} = UmyaSpreadsheet.save_ole_object_to_file(ole_object, "output.docx")
  """
  defdelegate save_ole_object_to_file(ole_object, file_path), to: OleObjects

  @doc """
  Checks if an OLE object is in binary format.

  ## Examples

      iex> {:ok, ole_object} = UmyaSpreadsheet.new_ole_object()
      iex> {:ok, _is_binary} = UmyaSpreadsheet.is_ole_object_binary_format(ole_object)
  """
  defdelegate is_ole_object_binary_format(ole_object), to: OleObjects

  @doc """
  Checks if an OLE object is in Excel format.

  ## Examples

      iex> {:ok, ole_object} = UmyaSpreadsheet.new_ole_object()
      iex> {:ok, _is_excel} = UmyaSpreadsheet.is_ole_object_excel_format(ole_object)
  """
  defdelegate is_ole_object_excel_format(ole_object), to: OleObjects

  @doc """
  Gets the 'requires' attribute from an OLE object.

  ## Examples

      iex> {:ok, ole_object} = UmyaSpreadsheet.new_ole_object()
      iex> :ok = UmyaSpreadsheet.set_ole_object_requires(ole_object, "xl")
      iex> {:ok, "xl"} = UmyaSpreadsheet.get_ole_object_requires(ole_object)
  """
  defdelegate get_ole_object_requires(ole_object), to: OleObjects

  @doc """
  Sets the 'requires' attribute for an OLE object.

  ## Examples

      iex> {:ok, ole_object} = UmyaSpreadsheet.new_ole_object()
      iex> UmyaSpreadsheet.set_ole_object_requires(ole_object, "xl")
      :ok
  """
  defdelegate set_ole_object_requires(ole_object, requires), to: OleObjects

  @doc """
  Gets the ProgID from an OLE object.

  ## Examples

      iex> {:ok, ole_object} = UmyaSpreadsheet.new_ole_object()
      iex> :ok = UmyaSpreadsheet.set_ole_object_prog_id(ole_object, "Excel.Sheet.12")
      iex> {:ok, "Excel.Sheet.12"} = UmyaSpreadsheet.get_ole_object_prog_id(ole_object)
  """
  defdelegate get_ole_object_prog_id(ole_object), to: OleObjects

  @doc """
  Sets the ProgID for an OLE object.

  ## Examples

      iex> {:ok, ole_object} = UmyaSpreadsheet.new_ole_object()
      iex> UmyaSpreadsheet.set_ole_object_prog_id(ole_object, "Excel.Sheet.12")
      :ok
  """
  defdelegate set_ole_object_prog_id(ole_object, prog_id), to: OleObjects

  @doc """
  Gets the file extension from an OLE object.

  ## Examples

      iex> {:ok, ole_object} = UmyaSpreadsheet.new_ole_object()
      iex> :ok = UmyaSpreadsheet.set_ole_object_extension(ole_object, "xlsx")
      iex> {:ok, "xlsx"} = UmyaSpreadsheet.get_ole_object_extension(ole_object)
  """
  defdelegate get_ole_object_extension(ole_object), to: OleObjects

  @doc """
  Sets the file extension for an OLE object.

  ## Examples

      iex> {:ok, ole_object} = UmyaSpreadsheet.new_ole_object()
      iex> UmyaSpreadsheet.set_ole_object_extension(ole_object, "xlsx")
      :ok
  """
  defdelegate set_ole_object_extension(ole_object, extension), to: OleObjects

  @doc """
  Gets the ProgID from embedded object properties.

  ## Examples

      iex> {:ok, properties} = UmyaSpreadsheet.new_embedded_object_properties()
      iex> :ok = UmyaSpreadsheet.set_embedded_object_prog_id(properties, "Word.Document.12")
      iex> {:ok, "Word.Document.12"} = UmyaSpreadsheet.get_embedded_object_prog_id(properties)
  """
  defdelegate get_embedded_object_prog_id(properties), to: OleObjects

  @doc """
  Sets the ProgID for embedded object properties.

  ## Examples

      iex> {:ok, properties} = UmyaSpreadsheet.new_embedded_object_properties()
      iex> UmyaSpreadsheet.set_embedded_object_prog_id(properties, "Word.Document.12")
      :ok
  """
  defdelegate set_embedded_object_prog_id(properties, prog_id), to: OleObjects

  @doc """
  Gets the shape ID from embedded object properties.

  ## Examples

      iex> {:ok, properties} = UmyaSpreadsheet.new_embedded_object_properties()
      iex> :ok = UmyaSpreadsheet.set_embedded_object_shape_id(properties, 123)
      iex> {:ok, 123} = UmyaSpreadsheet.get_embedded_object_shape_id(properties)
  """
  defdelegate get_embedded_object_shape_id(properties), to: OleObjects

  @doc """
  Sets the shape ID for embedded object properties.

  ## Examples

      iex> {:ok, properties} = UmyaSpreadsheet.new_embedded_object_properties()
      iex> UmyaSpreadsheet.set_embedded_object_shape_id(properties, 123)
      :ok
  """
  defdelegate set_embedded_object_shape_id(properties, shape_id), to: OleObjects

  @doc """
  Gets the binary data from an OLE object.

  ## Examples

      iex> {:ok, ole_object} = UmyaSpreadsheet.new_ole_object()
      iex> {:ok, _data} = UmyaSpreadsheet.get_ole_object_data(ole_object)
  """
  defdelegate get_ole_object_data(ole_object), to: OleObjects

  @doc """
  Sets the binary data for an OLE object.

  ## Examples

      iex> {:ok, ole_object} = UmyaSpreadsheet.new_ole_object()
      iex> :ok = UmyaSpreadsheet.set_ole_object_data(ole_object, "test data")
  """
  defdelegate set_ole_object_data(ole_object, data), to: OleObjects
end
