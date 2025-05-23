defmodule UmyaSpreadsheet.BackgroundFunctions do
  @moduledoc """
  Functions for manipulating cell background colors in a spreadsheet.
  """

  alias UmyaSpreadsheet.Spreadsheet
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
    case UmyaNative.set_background_color(ref, sheet_name, cell_address, color) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end
end
