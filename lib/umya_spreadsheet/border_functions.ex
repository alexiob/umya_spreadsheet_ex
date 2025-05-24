defmodule UmyaSpreadsheet.BorderFunctions do
  @moduledoc """
  Functions for manipulating cell borders in a spreadsheet.
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaSpreadsheet.ErrorHandling
  alias UmyaNative

  @doc """
  Sets the border style for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `border_position` - The border position ("top", "right", "bottom", "left", "diagonal", "outline", "all")
  - `border_style` - The border style ("thin", "medium", "thick", "dashed", "dotted", "double", "none")

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      # Set top border to thick
      :ok = UmyaSpreadsheet.BorderFunctions.set_border_style(spreadsheet, "Sheet1", "A1", "top", "thick")

      # Set all borders to thin
      :ok = UmyaSpreadsheet.BorderFunctions.set_border_style(spreadsheet, "Sheet1", "B2", "all", "thin")
  """
  def set_border_style(%Spreadsheet{reference: ref}, sheet_name, cell_address, border_position, border_style) do
    UmyaNative.set_border_style(ref, sheet_name, cell_address, border_position, border_style)
    |> ErrorHandling.standardize_result()
  end
end
