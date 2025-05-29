defmodule UmyaSpreadsheet.SheetViewFunctions do
  @moduledoc """
  Functions for configuring how sheets are displayed in the Excel application.
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaNative

  @doc """
  Sets whether to show gridlines in a sheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `show_gridlines` - Boolean indicating whether to show gridlines

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.SheetViewFunctions.set_show_grid_lines(spreadsheet, "Sheet1", false)
  """
  def set_show_grid_lines(%Spreadsheet{reference: ref}, sheet_name, show_gridlines) do
    case UmyaNative.set_show_grid_lines(ref, sheet_name, show_gridlines) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets whether a sheet tab is selected (active) when the workbook is opened.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `selected` - Boolean indicating whether the sheet tab should be selected

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.SheetViewFunctions.set_tab_selected(spreadsheet, "Sheet1", true)
  """
  def set_tab_selected(%Spreadsheet{reference: ref}, sheet_name, selected) do
    case UmyaNative.set_tab_selected(ref, sheet_name, selected) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets the top-left cell in the sheet view (the first visible cell).

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5") to set as the top-left cell

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.SheetViewFunctions.set_top_left_cell(spreadsheet, "Sheet1", "B5")
  """
  def set_top_left_cell(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    case UmyaNative.set_top_left_cell(ref, sheet_name, cell_address) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets the zoom scale for a sheet view.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `scale` - The zoom scale percentage (e.g., 100 for 100%, 150 for 150%)

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.SheetViewFunctions.set_zoom_scale(spreadsheet, "Sheet1", 150)
  """
  def set_zoom_scale(%Spreadsheet{reference: ref}, sheet_name, scale) do
    case UmyaNative.set_zoom_scale(ref, sheet_name, scale) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Freezes panes at the specified row and column.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `rows` - The number of rows to freeze
  - `cols` - The number of columns to freeze

  ## Examples

      # Freeze the top row
      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.SheetViewFunctions.freeze_panes(spreadsheet, "Sheet1", 1, 0)

      # Freeze the left column
      :ok = UmyaSpreadsheet.SheetViewFunctions.freeze_panes(spreadsheet, "Sheet1", 0, 1)

      # Freeze both top row and left column
      :ok = UmyaSpreadsheet.SheetViewFunctions.freeze_panes(spreadsheet, "Sheet1", 1, 1)
  """
  def freeze_panes(%Spreadsheet{reference: ref}, sheet_name, rows, cols) do
    case UmyaNative.freeze_panes(ref, sheet_name, rows, cols) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Splits panes at the specified row and column.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `height` - The height of the split in pixels
  - `width` - The width of the split in pixels

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.SheetViewFunctions.split_panes(spreadsheet, "Sheet1", 2000, 2000)
  """
  def split_panes(%Spreadsheet{reference: ref}, sheet_name, height, width) do
    case UmyaNative.split_panes(ref, sheet_name, height, width) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets the tab color for a worksheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `color` - The color code (e.g., "#FF0000" for red)

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      # Set sheet tab to red
      :ok = UmyaSpreadsheet.SheetViewFunctions.set_tab_color(spreadsheet, "Sheet1", "#FF0000")
  """
  def set_tab_color(%Spreadsheet{reference: ref}, sheet_name, color) do
    case UmyaNative.set_tab_color(ref, sheet_name, color) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets the sheet view type.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `view_type` - The type of view: "normal", "page_layout", or "page_break_preview"

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      # Set to normal view
      :ok = UmyaSpreadsheet.SheetViewFunctions.set_sheet_view(spreadsheet, "Sheet1", "normal")

      # Set to page layout view
      :ok = UmyaSpreadsheet.SheetViewFunctions.set_sheet_view(spreadsheet, "Sheet1", "page_layout")

      # Set to page break preview
      :ok = UmyaSpreadsheet.SheetViewFunctions.set_sheet_view(spreadsheet, "Sheet1", "page_break_preview")
  """
  def set_sheet_view(%Spreadsheet{reference: ref}, sheet_name, view_type) do
    case UmyaNative.set_sheet_view(ref, sheet_name, view_type) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets the zoom scale for the normal view.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `scale` - The zoom scale percentage (e.g., 100 for 100%, 150 for 150%)

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.SheetViewFunctions.set_zoom_scale_normal(spreadsheet, "Sheet1", 75)
  """
  def set_zoom_scale_normal(%Spreadsheet{reference: ref}, sheet_name, scale) do
    case UmyaNative.set_zoom_scale_normal(ref, sheet_name, scale) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets the zoom scale for the page layout view.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `scale` - The zoom scale percentage (e.g., 100 for 100%, 150 for 150%)

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.SheetViewFunctions.set_zoom_scale_page_layout(spreadsheet, "Sheet1", 120)
  """
  def set_zoom_scale_page_layout(%Spreadsheet{reference: ref}, sheet_name, scale) do
    case UmyaNative.set_zoom_scale_page_layout(ref, sheet_name, scale) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets the zoom scale for the page break preview view.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `scale` - The zoom scale percentage (e.g., 100 for 100%, 150 for 150%)

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.SheetViewFunctions.set_zoom_scale_page_break(spreadsheet, "Sheet1", 80)
  """
  def set_zoom_scale_page_break(%Spreadsheet{reference: ref}, sheet_name, scale) do
    case UmyaNative.set_zoom_scale_page_break(ref, sheet_name, scale) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets the active selection in a worksheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `active_cell` - The active cell address (e.g., "A1", "B5")
  - `sqref` - The selection range reference (e.g., "A1:B5")

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.SheetViewFunctions.set_selection(spreadsheet, "Sheet1", "C3", "C3:D5")
  """
  def set_selection(%Spreadsheet{reference: ref}, sheet_name, active_cell, sqref) do
    case UmyaNative.set_selection(ref, sheet_name, active_cell, sqref) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Gets whether gridlines are shown in a sheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet

  ## Returns

  - `{:ok, boolean}` where `boolean` indicates whether gridlines are shown
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, show_gridlines} = UmyaSpreadsheet.SheetViewFunctions.get_show_grid_lines(spreadsheet, "Sheet1")
      # show_gridlines = true
  """
  def get_show_grid_lines(%Spreadsheet{reference: ref}, sheet_name) do
    UmyaNative.get_show_grid_lines(ref, sheet_name)
  end

  @doc """
  Gets the zoom scale for a sheet view.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet

  ## Returns

  - `{:ok, integer}` where `integer` is the zoom scale percentage
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, zoom_scale} = UmyaSpreadsheet.SheetViewFunctions.get_zoom_scale(spreadsheet, "Sheet1")
      # zoom_scale = 100
  """
  def get_zoom_scale(%Spreadsheet{reference: ref}, sheet_name) do
    UmyaNative.get_zoom_scale(ref, sheet_name)
  end

  @doc """
  Gets the tab color for a worksheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet

  ## Returns

  - `{:ok, color}` where `color` is a string with the color code (e.g., "#FF0000" for red)
    or an empty string if no color is set
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, color} = UmyaSpreadsheet.SheetViewFunctions.get_tab_color(spreadsheet, "Sheet1")
      # color = "#FF0000"
  """
  def get_tab_color(%Spreadsheet{reference: ref}, sheet_name) do
    UmyaNative.get_tab_color(ref, sheet_name)
  end

  @doc """
  Gets the sheet view type.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet

  ## Returns

  - `{:ok, view_type}` where `view_type` is one of "normal", "page_layout", or "page_break_preview"
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, view_type} = UmyaSpreadsheet.SheetViewFunctions.get_sheet_view(spreadsheet, "Sheet1")
      # view_type = "normal"
  """
  def get_sheet_view(%Spreadsheet{reference: ref}, sheet_name) do
    UmyaNative.get_sheet_view(ref, sheet_name)
  end

  @doc """
  Gets the active selection in a worksheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet

  ## Returns

  - `{:ok, selection}` where `selection` is a map with "active_cell" and "sqref" keys
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, selection} = UmyaSpreadsheet.SheetViewFunctions.get_selection(spreadsheet, "Sheet1")
      # selection = %{"active_cell" => "A1", "sqref" => "A1:B5"}
  """
  def get_selection(%Spreadsheet{reference: ref}, sheet_name) do
    UmyaNative.get_selection(ref, sheet_name)
  end
end
