defmodule UmyaSpreadsheet.BackwardCompatibility do
  @moduledoc """
  This module provides backward compatibility for function names that have changed over time.
  It's intended to ease the transition for users upgrading from previous versions.
  """
  
  alias UmyaSpreadsheet.Spreadsheet
  
  @doc """
  Backward compatibility for set_show_gridlines which is now set_show_grid_lines.
  
  @deprecated Use set_show_grid_lines/3 instead
  """
  def set_show_gridlines(%Spreadsheet{} = spreadsheet, sheet_name, show_gridlines) do
    UmyaSpreadsheet.set_show_grid_lines(spreadsheet, sheet_name, show_gridlines)
  end
end
