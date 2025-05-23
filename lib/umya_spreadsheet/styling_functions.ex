defmodule UmyaSpreadsheet.StylingFunctions do
  @moduledoc """
  Functions for styling and formatting cells, columns, and rows.
  """
  
  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaNative

  @doc """
  Copies styling from one column to another within the same sheet.
  
  ## Parameters
  
  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet containing both columns
  - `source_column` - The index of the source column (1-based)
  - `target_column` - The index of the target column (1-based)
  - `start_row` - Optional starting row (1-based, defaults to 1)
  - `end_row` - Optional ending row (1-based, defaults to include all rows)
  
  ## Returns
  
  - `:ok` on success
  - `{:error, reason}` on failure
  
  ## Examples
  
      # Copy all styling from column A to column B
      :ok = UmyaSpreadsheet.copy_column_styling(spreadsheet, "Sheet1", 1, 2)
      
      # Copy styling from column A to column B, but only for rows 3-10
      :ok = UmyaSpreadsheet.copy_column_styling(spreadsheet, "Sheet1", 1, 2, 3, 10)
  """
  def copy_column_styling(%Spreadsheet{reference: ref}, sheet_name, source_column, target_column, start_row \\ nil, end_row \\ nil) do
    case UmyaNative.copy_column_styling(ref, sheet_name, source_column, target_column, start_row, end_row) do
      {:ok, :ok} -> :ok
      result -> result
    end
  end
end
