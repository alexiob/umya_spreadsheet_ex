defmodule UmyaSpreadsheet.FontFunctions do
  @moduledoc """
  Functions for manipulating font styles in a spreadsheet.
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaNative

  @doc """
  Sets the font color for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `color` - The color code (e.g., "#FF0000" for red)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      # Set text to red
      :ok = UmyaSpreadsheet.FontFunctions.set_font_color(spreadsheet, "Sheet1", "A1", "#FF0000")
  """
  def set_font_color(%Spreadsheet{reference: ref}, sheet_name, cell_address, color) do
    case UmyaNative.set_font_color(ref, sheet_name, cell_address, color) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets the font size for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `size` - The font size in points

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.FontFunctions.set_font_size(spreadsheet, "Sheet1", "A1", 14)
  """
  def set_font_size(%Spreadsheet{reference: ref}, sheet_name, cell_address, size) do
    case UmyaNative.set_font_size(ref, sheet_name, cell_address, size) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets the font weight (bold) for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `is_bold` - Boolean indicating whether the font should be bold

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.FontFunctions.set_font_bold(spreadsheet, "Sheet1", "A1", true)
  """
  def set_font_bold(%Spreadsheet{reference: ref}, sheet_name, cell_address, is_bold) do
    case UmyaNative.set_font_bold(ref, sheet_name, cell_address, is_bold) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets the font name for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `font_name` - The name of the font

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.FontFunctions.set_font_name(spreadsheet, "Sheet1", "A1", "Arial")
  """
  def set_font_name(%Spreadsheet{reference: ref}, sheet_name, cell_address, font_name) do
    case UmyaNative.set_font_name(ref, sheet_name, cell_address, font_name) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets the font style to italic for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `is_italic` - Boolean indicating whether the font should be italic

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.FontFunctions.set_font_italic(spreadsheet, "Sheet1", "A1", true)
  """
  def set_font_italic(%Spreadsheet{reference: ref}, sheet_name, cell_address, is_italic) do
    case UmyaNative.set_font_italic(ref, sheet_name, cell_address, is_italic) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets the font underline style for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `underline_style` - The underline style ("single", "double", "none")

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.FontFunctions.set_font_underline(spreadsheet, "Sheet1", "A1", "single")
  """
  def set_font_underline(%Spreadsheet{reference: ref}, sheet_name, cell_address, underline_style) do
    case UmyaNative.set_font_underline(ref, sheet_name, cell_address, underline_style) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets the font strikethrough property for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `is_strikethrough` - Boolean indicating whether the font should have strikethrough

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.FontFunctions.set_font_strikethrough(spreadsheet, "Sheet1", "A1", true)
  """
  def set_font_strikethrough(%Spreadsheet{reference: ref}, sheet_name, cell_address, is_strikethrough) do
    case UmyaNative.set_font_strikethrough(ref, sheet_name, cell_address, is_strikethrough) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end
end
