defmodule UmyaSpreadsheet.CellFunctions do
  @moduledoc """
  Functions for manipulating individual cells in a spreadsheet.
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaSpreadsheet.ErrorHandling
  alias UmyaNative

  @doc """
  Gets the value of a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - The cell value as a string
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      "Hello" = UmyaSpreadsheet.CellFunctions.get_cell_value(spreadsheet, "Sheet1", "A1")
  """
  def get_cell_value(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    UmyaNative.get_cell_value(ref, sheet_name, cell_address)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the formatted value of a cell (as displayed in Excel).

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - The formatted cell value as a string
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      "12/31/2023" = UmyaSpreadsheet.CellFunctions.get_formatted_value(spreadsheet, "Sheet1", "A1")
  """
  def get_formatted_value(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    UmyaNative.get_formatted_value(ref, sheet_name, cell_address)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the value of a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `value` - The value to set

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.CellFunctions.set_cell_value(spreadsheet, "Sheet1", "A1", "Hello")
  """
  def set_cell_value(%Spreadsheet{reference: ref}, sheet_name, cell_address, value) do
    # Convert value to string for compatibility with the Rust NIF
    string_value =
      cond do
        is_binary(value) -> value
        is_integer(value) -> Integer.to_string(value)
        is_float(value) -> Float.to_string(value)
        is_boolean(value) -> to_string(value)
        is_nil(value) -> ""
        is_atom(value) -> Atom.to_string(value)
        true -> inspect(value)
      end

    UmyaNative.set_cell_value(ref, sheet_name, cell_address, string_value)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Removes a cell from the spreadsheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.CellFunctions.remove_cell(spreadsheet, "Sheet1", "A1")
  """
  def remove_cell(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    UmyaNative.remove_cell(ref, sheet_name, cell_address)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the number format for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `format_code` - The format code (e.g., "0.00", "m/d/yyyy")

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      # Set number format to 2 decimal places
      :ok = UmyaSpreadsheet.CellFunctions.set_number_format(spreadsheet, "Sheet1", "A1", "0.00")
      # Set date format
      :ok = UmyaSpreadsheet.CellFunctions.set_number_format(spreadsheet, "Sheet1", "B1", "m/d/yyyy")
  """
  def set_number_format(%Spreadsheet{reference: ref}, sheet_name, cell_address, format_code) do
    UmyaNative.set_number_format(ref, sheet_name, cell_address, format_code)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets text wrapping for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `wrap` - Boolean indicating whether to wrap text

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.CellFunctions.set_wrap_text(spreadsheet, "Sheet1", "A1", true)
  """
  def set_wrap_text(%Spreadsheet{reference: ref}, sheet_name, cell_address, wrap) do
    UmyaNative.set_wrap_text(ref, sheet_name, cell_address, wrap)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the cell alignment.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `horizontal` - Horizontal alignment ("left", "center", "right", "justify")
  - `vertical` - Vertical alignment ("top", "center", "bottom")

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.CellFunctions.set_cell_alignment(spreadsheet, "Sheet1", "A1", "center", "center")
  """
  def set_cell_alignment(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_address,
        horizontal,
        vertical
      ) do
    UmyaNative.set_cell_alignment(ref, sheet_name, cell_address, horizontal, vertical)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the text rotation angle for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `angle` - The rotation angle in degrees

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.CellFunctions.set_cell_rotation(spreadsheet, "Sheet1", "A1", 45)
  """
  def set_cell_rotation(%Spreadsheet{reference: ref}, sheet_name, cell_address, angle) do
    UmyaNative.set_cell_rotation(ref, sheet_name, cell_address, angle)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the text indentation level for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `level` - The indentation level

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.CellFunctions.set_cell_indent(spreadsheet, "Sheet1", "A1", 2)
  """
  def set_cell_indent(%Spreadsheet{reference: ref}, sheet_name, cell_address, level) do
    UmyaNative.set_cell_indent(ref, sheet_name, cell_address, level)
    |> ErrorHandling.standardize_result()
  end

  # ============================================================================
  # CELL FORMATTING GETTER FUNCTIONS
  # ============================================================================

  @doc """
  Gets the horizontal alignment of a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - `{:ok, alignment}` where alignment is "general", "left", "center", "right", etc.
  - `{:error, reason}` on failure
  """
  def get_cell_horizontal_alignment(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    UmyaNative.get_cell_horizontal_alignment(ref, sheet_name, cell_address)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the vertical alignment of a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - `{:ok, alignment}` where alignment is "top", "center", "bottom", etc.
  - `{:error, reason}` on failure
  """
  def get_cell_vertical_alignment(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    UmyaNative.get_cell_vertical_alignment(ref, sheet_name, cell_address)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the wrap text setting of a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - `{:ok, wrap}` where wrap is a boolean indicating if text wrapping is enabled
  - `{:error, reason}` on failure
  """
  def get_cell_wrap_text(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    UmyaNative.get_cell_wrap_text(ref, sheet_name, cell_address)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the text rotation of a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - `{:ok, angle}` where angle is the rotation angle in degrees
  - `{:error, reason}` on failure
  """
  def get_cell_text_rotation(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    UmyaNative.get_cell_text_rotation(ref, sheet_name, cell_address)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the border style of a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `border_position` - "top", "bottom", "left", "right", or "diagonal"

  ## Returns

  - `{:ok, style}` where style is the border style (e.g., "solid", "dashed", "dotted")
  - `{:error, reason}` on failure
  """
  def get_border_style(%Spreadsheet{reference: ref}, sheet_name, cell_address, border_position) do
    UmyaNative.get_border_style(ref, sheet_name, cell_address, border_position)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the border color of a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `border_position` - "top", "bottom", "left", "right", or "diagonal"

  ## Returns

  - `{:ok, color}` where color is the border color in hex format (e.g., "#FF0000")
  - `{:error, reason}` on failure
  """
  def get_border_color(%Spreadsheet{reference: ref}, sheet_name, cell_address, border_position) do
    UmyaNative.get_border_color(ref, sheet_name, cell_address, border_position)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the background color of a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - `{:ok, color}` where color is the background color in hex format (e.g., "#FFFFFF")
  - `{:error, reason}` on failure
  """
  def get_cell_background_color(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    UmyaNative.get_cell_background_color(ref, sheet_name, cell_address)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the foreground color of a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - `{:ok, color}` where color is the foreground color in hex format (e.g., "#000000")
  - `{:error, reason}` on failure
  """
  def get_cell_foreground_color(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    UmyaNative.get_cell_foreground_color(ref, sheet_name, cell_address)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the pattern type of a cell's fill.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - `{:ok, pattern_type}` where pattern_type is the pattern type (e.g., "solid", "none")
  - `{:error, reason}` on failure
  """
  def get_cell_pattern_type(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    UmyaNative.get_cell_pattern_type(ref, sheet_name, cell_address)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the number format ID of a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - `{:ok, format_id}` where format_id is the number format ID
  - `{:error, reason}` on failure
  """
  def get_cell_number_format_id(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    UmyaNative.get_cell_number_format_id(ref, sheet_name, cell_address)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the number format code of a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - `{:ok, format_code}` where format_code is the number format code (e.g., "0.00", "m/d/yyyy")
  - `{:error, reason}` on failure
  """
  def get_cell_format_code(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    UmyaNative.get_cell_format_code(ref, sheet_name, cell_address)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the locked status of a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - `{:ok, locked}` where locked is a boolean indicating if the cell is locked
  - `{:error, reason}` on failure
  """
  def get_cell_locked(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    UmyaNative.get_cell_locked(ref, sheet_name, cell_address)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the hidden status of a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - `{:ok, hidden}` where hidden is a boolean indicating if the cell is hidden
  - `{:error, reason}` on failure
  """
  def get_cell_hidden(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    UmyaNative.get_cell_hidden(ref, sheet_name, cell_address)
    |> ErrorHandling.standardize_result()
  end
end
