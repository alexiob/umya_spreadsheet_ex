defmodule UmyaSpreadsheet.CellFunctions do
  @moduledoc """
  Functions for manipulating individual cells in a spreadsheet.
  """
  
  alias UmyaSpreadsheet.Spreadsheet
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
    case UmyaNative.get_cell_value(ref, sheet_name, cell_address) do
      {:error, reason} -> {:error, reason}
      value -> value
    end
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
    case UmyaNative.get_formatted_value(ref, sheet_name, cell_address) do
      {:error, reason} -> {:error, reason}
      value -> value
    end
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
    string_value = cond do
      is_binary(value) -> value
      is_integer(value) -> Integer.to_string(value)
      is_float(value) -> Float.to_string(value)
      is_boolean(value) -> to_string(value)
      is_nil(value) -> ""
      is_atom(value) -> Atom.to_string(value)
      true -> inspect(value)
    end
    
    case UmyaNative.set_cell_value(ref, sheet_name, cell_address, string_value) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
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
    case UmyaNative.remove_cell(ref, sheet_name, cell_address) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
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
    case UmyaNative.set_number_format(ref, sheet_name, cell_address, format_code) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
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
    case UmyaNative.set_wrap_text(ref, sheet_name, cell_address, wrap) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
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
  def set_cell_alignment(%Spreadsheet{reference: ref}, sheet_name, cell_address, horizontal, vertical) do
    case UmyaNative.set_cell_alignment(ref, sheet_name, cell_address, horizontal, vertical) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
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
    case UmyaNative.set_cell_rotation(ref, sheet_name, cell_address, angle) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
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
    case UmyaNative.set_cell_indent(ref, sheet_name, cell_address, level) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end
end
