defmodule UmyaSpreadsheet.WorkbookViewFunctions do
  @moduledoc """
  Functions for configuring how the workbook is displayed in the Excel application.
  These settings affect the overall workbook window rather than individual sheet views.

  This module provides functions to:
  * Get and set the active tab (worksheet) when the workbook is opened
  * Get and set the position and size of the Excel application window
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaSpreadsheet.ErrorHandling
  alias UmyaNative

  @doc """
  Sets which tab (worksheet) is active when the workbook is opened.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `tab_index` - The zero-based index of the tab to make active

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      # Make the second tab active when opening the file
      :ok = UmyaSpreadsheet.WorkbookViewFunctions.set_active_tab(spreadsheet, 1)

      # Make the first tab active
      :ok = UmyaSpreadsheet.WorkbookViewFunctions.set_active_tab(spreadsheet, 0)
  """
  def set_active_tab(%Spreadsheet{reference: ref}, tab_index) do
    UmyaNative.set_active_tab(ref, tab_index)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the position and size of the workbook window when opened in Excel.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `x_position` - The horizontal position of the window in pixels
  - `y_position` - The vertical position of the window in pixels
  - `window_width` - The width of the window in pixels
  - `window_height` - The height of the window in pixels

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      # Set the workbook to open at position (100, 50) with a size of 800x600
      :ok = UmyaSpreadsheet.WorkbookViewFunctions.set_workbook_window_position(spreadsheet, 100, 50, 800, 600)
  """
  def set_workbook_window_position(
        %Spreadsheet{reference: ref},
        x_position,
        y_position,
        window_width,
        window_height
      ) do
    UmyaNative.set_workbook_window_position(
      ref,
      x_position,
      y_position,
      window_width,
      window_height
    )
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the index of the active tab (worksheet) when the workbook is opened.

  Tab indices are zero-based, meaning the first tab has index 0, the second has index 1, etc.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct

  ## Returns

  - `{:ok, tab_index}` where tab_index is the zero-based index of the active tab
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, active_tab} = UmyaSpreadsheet.WorkbookViewFunctions.get_active_tab(spreadsheet)
      # active_tab = 2 (third worksheet is active)
  """
  def get_active_tab(%Spreadsheet{reference: ref}) do
    UmyaNative.get_active_tab(ref)
  end

  @doc """
  Gets the position and size settings for the workbook window in Excel.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct

  ## Returns

  - `{:ok, window_info}` where window_info is a map containing position and size information
    with keys `:x_position`, `:y_position`, `:width`, and `:height`
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, window_info} = UmyaSpreadsheet.WorkbookViewFunctions.get_workbook_window_position(spreadsheet)
      # window_info = %{"x_position" => "240", "y_position" => "105", "width" => "14805", "height" => "8010"}
  """
  def get_workbook_window_position(%Spreadsheet{reference: ref}) do
    UmyaNative.get_workbook_window_position(ref)
  end
end
