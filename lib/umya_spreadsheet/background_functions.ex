defmodule UmyaSpreadsheet.BackgroundFunctions do
  @moduledoc """
  Functions for manipulating and inspecting cell background colors in a spreadsheet.
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaSpreadsheet.ErrorHandling
  alias UmyaNative

  @doc """
  Sets the background color of a cell.

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
      # Set cell background to light blue
      :ok = UmyaSpreadsheet.BackgroundFunctions.set_background_color(spreadsheet, "Sheet1", "A1", "#CCECFF")
  """
  def set_background_color(%Spreadsheet{reference: ref}, sheet_name, cell_address, color) do
    UmyaNative.set_background_color(ref, sheet_name, cell_address, color)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the background color of a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - `{:ok, color}` on success where color is a hex color code (e.g., "FFFFFF00")
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, color} = UmyaSpreadsheet.BackgroundFunctions.get_cell_background_color(spreadsheet, "Sheet1", "A1")
      # => "FFFFFF00" (yellow background)
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

  - `{:ok, color}` on success where color is a hex color code (e.g., "FFFFFFFF")
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, color} = UmyaSpreadsheet.BackgroundFunctions.get_cell_foreground_color(spreadsheet, "Sheet1", "A1")
      # => "FFFFFFFF" (white foreground)
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

  - `{:ok, pattern_type}` on success where pattern_type is a string (e.g., "solid", "none")
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, pattern} = UmyaSpreadsheet.BackgroundFunctions.get_cell_pattern_type(spreadsheet, "Sheet1", "A1")
      # => "solid"
  """
  def get_cell_pattern_type(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    UmyaNative.get_cell_pattern_type(ref, sheet_name, cell_address)
    |> ErrorHandling.standardize_result()
  end
end
