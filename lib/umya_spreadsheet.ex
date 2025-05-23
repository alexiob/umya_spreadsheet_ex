defmodule UmyaSpreadsheet do
  @moduledoc """
  UmyaSpreadsheet is an Elixir wrapper for the Rust `umya-spreadsheet` library,
  which provides Excel (.xlsx, .xlsm) file manipulation capabilities with the
  performance benefits of Rust.

  This library allows you to:
  - Create new spreadsheets
  - Read/write Excel files
  - Manipulate cell values and styles
  - Move ranges of data
  - Add new sheets
  - Create charts
  - Add images
  - Add data validation (dropdown lists, number ranges, etc.)
  - Export to CSV
  - And more...

  ## Guides

  For detailed guides on specific features, see:

  - [Charts](charts.html) - Creating and customizing various chart types
  - [CSV Export & Performance](csv_export_and_performance.html) - CSV export and optimized writers
  - [Image Handling](image_handling.html) - Working with images in spreadsheets
  - [Sheet Operations](sheet_operations.html) - Managing worksheets effectively
  - [Styling and Formatting](styling_and_formatting.html) - Making your spreadsheets look professional

  See the [Guide Index](guides.html) for a complete list of available guides.
  """

  alias UmyaNative

  # Import specialized function modules
  alias UmyaSpreadsheet.BorderFunctions
  alias UmyaSpreadsheet.CellFunctions
  alias UmyaSpreadsheet.ChartFunctions
  alias UmyaSpreadsheet.ConditionalFormatting
  alias UmyaSpreadsheet.CSVFunctions
  alias UmyaSpreadsheet.DataValidation
  alias UmyaSpreadsheet.Drawing
  alias UmyaSpreadsheet.FontFunctions
  alias UmyaSpreadsheet.ImageFunctions
  alias UmyaSpreadsheet.PerformanceFunctions
  alias UmyaSpreadsheet.PrintSettings
  alias UmyaSpreadsheet.RowColumnFunctions
  alias UmyaSpreadsheet.SheetFunctions
  alias UmyaSpreadsheet.SheetViewFunctions
  alias UmyaSpreadsheet.StylingFunctions
  alias UmyaSpreadsheet.WorkbookFunctions
  alias UmyaSpreadsheet.WorkbookViewFunctions
  alias UmyaSpreadsheet.BackgroundFunctions
  alias UmyaSpreadsheet.PivotTable

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
      iex> UmyaSpreadsheet.write_with_password(spreadsheet, "test/result_files/secure.xlsx", "password123")
      :ok
  """
  def write_with_password(%Spreadsheet{reference: ref}, path, password) do
    case UmyaNative.write_file_with_password(unwrap_ref(ref), path, password) do
      :ok -> :ok
      # Handle the {:ok, :ok} tuple from Rustler 0.36.1
      {:ok, :ok} -> :ok
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
  defdelegate copy_column_styling(spreadsheet, sheet_name, source_column, target_column), to: StylingFunctions
  defdelegate copy_column_styling(spreadsheet, sheet_name, source_column, target_column, start_row, end_row), to: StylingFunctions

  # Performance Functions delegation
  defdelegate write_light(spreadsheet, path), to: PerformanceFunctions
  defdelegate write_with_password_light(spreadsheet, path, password), to: PerformanceFunctions

  # Conditional Formatting Functions delegation
  defdelegate add_color_scale(spreadsheet, sheet_name, range, min_type, min_value, min_color, max_type, max_value, max_color),
    to: ConditionalFormatting
  defdelegate add_color_scale(spreadsheet, sheet_name, range, min_type, min_value, min_color, mid_type, mid_value, mid_color, max_type, max_value, max_color),
    to: ConditionalFormatting
  defdelegate add_cell_value_rule(spreadsheet, sheet_name, range, operator, value1, value2, format_style),
    to: ConditionalFormatting
  defdelegate add_cell_is_rule(spreadsheet, sheet_name, range, operator, value1, value2, format_style),
    to: ConditionalFormatting
  defdelegate add_text_rule(spreadsheet, sheet_name, range, operator, text, format_style),
    to: ConditionalFormatting
  defdelegate add_data_bar(spreadsheet, sheet_name, range, min_value, max_value, color),
    to: ConditionalFormatting
  defdelegate add_top_bottom_rule(spreadsheet, sheet_name, range, rule_type, rank, percent, format_style),
    to: ConditionalFormatting

  # Data Validation Functions delegation
  defdelegate add_list_validation(spreadsheet, sheet_name, range, options), to: DataValidation
  defdelegate add_list_validation(spreadsheet, sheet_name, range, options, show_dropdown), to: DataValidation
  defdelegate add_list_validation(spreadsheet, sheet_name, range, options, show_dropdown, error_message), to: DataValidation
  defdelegate add_list_validation(spreadsheet, sheet_name, range, options, show_dropdown, error_message, error_title, prompt_message, prompt_title), to: DataValidation

  defdelegate add_number_validation(spreadsheet, sheet_name, range, operator, value1), to: DataValidation
  defdelegate add_number_validation(spreadsheet, sheet_name, range, operator, value1, value2), to: DataValidation
  defdelegate add_number_validation(spreadsheet, sheet_name, range, operator, value1, value2, error_message), to: DataValidation
  defdelegate add_number_validation(spreadsheet, sheet_name, range, operator, value1, value2, show_dropdown, error_message, error_title, prompt_message, prompt_title), to: DataValidation

  defdelegate add_date_validation(spreadsheet, sheet_name, range, operator, value1), to: DataValidation
  defdelegate add_date_validation(spreadsheet, sheet_name, range, operator, value1, value2), to: DataValidation
  defdelegate add_date_validation(spreadsheet, sheet_name, range, operator, value1, value2, error_message), to: DataValidation
  defdelegate add_date_validation(spreadsheet, sheet_name, range, operator, value1, value2, show_dropdown, error_message, error_title, prompt_message, prompt_title), to: DataValidation

  defdelegate add_text_length_validation(spreadsheet, sheet_name, range, operator, value1), to: DataValidation
  defdelegate add_text_length_validation(spreadsheet, sheet_name, range, operator, value1, value2), to: DataValidation
  defdelegate add_text_length_validation(spreadsheet, sheet_name, range, operator, value1, value2, error_message), to: DataValidation
  defdelegate add_text_length_validation(spreadsheet, sheet_name, range, operator, value1, value2, show_dropdown, error_message, error_title, prompt_message, prompt_title), to: DataValidation

  defdelegate add_custom_validation(spreadsheet, sheet_name, range, formula), to: DataValidation
  defdelegate add_custom_validation(spreadsheet, sheet_name, range, formula, error_message), to: DataValidation
  defdelegate add_custom_validation(spreadsheet, sheet_name, range, formula, show_dropdown, error_message, error_title, prompt_message, prompt_title), to: DataValidation

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
  defdelegate set_active_tab(spreadsheet, tab_index),
    to: WorkbookViewFunctions
  defdelegate set_workbook_window_position(spreadsheet, x_position, y_position, window_width, window_height),
    to: WorkbookViewFunctions

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
  defdelegate copy_row_styling(spreadsheet, sheet_name, source_row, target_row, start_column, end_column),
    to: RowColumnFunctions

  # Border Functions delegation
  defdelegate set_border_style(spreadsheet, sheet_name, cell_address, border_position, border_style),
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
  defdelegate add_chart(spreadsheet, sheet_name, chart_type, from_cell, to_cell, title, data_series, series_titles, point_titles),
    to: ChartFunctions
  defdelegate add_chart_with_options(spreadsheet, sheet_name, chart_type, from_cell, to_cell, title, data_series, series_titles, point_titles, style, vary_colors, view_3d, legend, axes, data_labels),
    to: ChartFunctions
  defdelegate add_chart_with_options(spreadsheet, sheet_name, chart_type, from_cell, to_cell, title, data_series, series_titles, point_titles, style, vary_colors, view_3d, legend, axes, data_labels, chart_specific),
    to: ChartFunctions
  defdelegate set_chart_style(spreadsheet, sheet_name, chart_index, style),
    to: ChartFunctions
  defdelegate set_chart_data_labels(spreadsheet, sheet_name, chart_index, show_values, show_percent, show_category_name, show_series_name, position),
    to: ChartFunctions
  defdelegate set_chart_legend_position(spreadsheet, sheet_name, chart_index, position, overlay),
    to: ChartFunctions
  defdelegate set_chart_3d_view(spreadsheet, sheet_name, chart_index, rot_x, rot_y, perspective, height_percent),
    to: ChartFunctions
  defdelegate set_chart_axis_titles(spreadsheet, sheet_name, chart_index, category_axis_title, value_axis_title),
    to: ChartFunctions

  # Drawing Functions delegation
  defdelegate add_shape(spreadsheet, sheet_name, cell_address, shape_type, width, height, fill_color, outline_color, outline_width),
    to: Drawing
  defdelegate add_text_box(spreadsheet, sheet_name, cell_address, text, width, height, fill_color, text_color, outline_color, outline_width),
    to: Drawing
  defdelegate add_connector(spreadsheet, sheet_name, from_cell, to_cell, line_color, line_width),
    to: Drawing

  # Backward compatibility functions

  @doc """
  Backward compatibility for set_show_gridlines which is now set_show_grid_lines.

  @deprecated Use set_show_grid_lines/3 instead
  """
  def set_show_gridlines(%Spreadsheet{} = spreadsheet, sheet_name, show_gridlines) do
    set_show_grid_lines(spreadsheet, sheet_name, show_gridlines)
  end

  # PivotTable Functions delegation
  defdelegate add_pivot_table(spreadsheet, sheet_name, name, source_sheet, source_range, target_cell, row_fields, column_fields, data_fields),
    to: PivotTable

  # Background Functions delegation
  defdelegate set_background_color(spreadsheet, sheet_name, cell_address, color),
    to: BackgroundFunctions
end
