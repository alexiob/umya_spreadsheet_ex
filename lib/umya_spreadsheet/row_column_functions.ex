defmodule UmyaSpreadsheet.RowColumnFunctions do
  @moduledoc """
  Functions for manipulating rows and columns in a spreadsheet.
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaNative

  @doc """
  Sets the height of a row.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `row_number` - The row number (1-based)
  - `height` - The height value (in points)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.RowColumnFunctions.set_row_height(spreadsheet, "Sheet1", 1, 20.5)
  """
  def set_row_height(%Spreadsheet{reference: ref}, sheet_name, row_number, height) do
    case UmyaNative.set_row_height(ref, sheet_name, row_number, height) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets styles for an entire row.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `row_number` - The row number (1-based)
  - `bg_color` - The background color code (e.g., "#FF0000" for red)
  - `font_color` - The font color code (e.g., "#FFFFFF" for white)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      # Set row background to red and text to white
      :ok = UmyaSpreadsheet.RowColumnFunctions.set_row_style(spreadsheet, "Sheet1", 1, "#FF0000", "#FFFFFF")
  """
  def set_row_style(%Spreadsheet{reference: ref}, sheet_name, row_number, bg_color, font_color) do
    case UmyaNative.set_row_style(ref, sheet_name, row_number, bg_color, font_color) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets the width of a column.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `column` - The column letter (e.g., "A", "B")
  - `width` - The width value

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.RowColumnFunctions.set_column_width(spreadsheet, "Sheet1", "A", 15.5)
  """
  def set_column_width(%Spreadsheet{reference: ref}, sheet_name, column, width) do
    case UmyaNative.set_column_width(ref, sheet_name, column, width) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets whether a column width should be automatically adjusted.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `column` - The column letter (e.g., "A", "B")
  - `auto_width` - Boolean indicating whether auto-width should be enabled

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.RowColumnFunctions.set_column_auto_width(spreadsheet, "Sheet1", "A", true)
  """
  def set_column_auto_width(%Spreadsheet{reference: ref}, sheet_name, column, auto_width) do
    case UmyaNative.set_column_auto_width(ref, sheet_name, column, auto_width) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Copies styling from one row to another.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `source_row` - The source row number (1-based)
  - `target_row` - The target row number (1-based)
  - `start_column` - Optional starting column index (1-based)
  - `end_column` - Optional ending column index (1-based)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      # Copy styling for all cells in row 1 to row 2
      :ok = UmyaSpreadsheet.RowColumnFunctions.copy_row_styling(spreadsheet, "Sheet1", 1, 2, nil, nil)

      # Copy styling only for columns A through C (columns 1-3)
      :ok = UmyaSpreadsheet.RowColumnFunctions.copy_row_styling(spreadsheet, "Sheet1", 1, 2, 1, 3)
  """
  def copy_row_styling(
        %Spreadsheet{reference: ref},
        sheet_name,
        source_row,
        target_row,
        start_column \\ nil,
        end_column \\ nil
      ) do
    case UmyaNative.copy_row_styling(
           ref,
           sheet_name,
           source_row,
           target_row,
           start_column,
           end_column
         ) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Copies styling from one column to another.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `source_column` - The source column index (1-based)
  - `target_column` - The target column index (1-based)
  - `start_row` - Optional starting row index (1-based)
  - `end_row` - Optional ending row index (1-based)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      # Copy styling for all cells in column A to column B
      :ok = UmyaSpreadsheet.RowColumnFunctions.copy_column_styling(spreadsheet, "Sheet1", 1, 2, nil, nil)

      # Copy styling only for rows 1 through 10
      :ok = UmyaSpreadsheet.RowColumnFunctions.copy_column_styling(spreadsheet, "Sheet1", 1, 2, 1, 10)
  """
  def copy_column_styling(
        %Spreadsheet{reference: ref},
        sheet_name,
        source_column,
        target_column,
        start_row \\ nil,
        end_row \\ nil
      ) do
    case UmyaNative.copy_column_styling(
           ref,
           sheet_name,
           source_column,
           target_column,
           start_row,
           end_row
         ) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Gets the width of a column.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `column` - The column letter (e.g., "A", "B")

  ## Returns

  - `{:ok, width}` on success where width is a float
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, width} = UmyaSpreadsheet.RowColumnFunctions.get_column_width(spreadsheet, "Sheet1", "A")
  """
  def get_column_width(%Spreadsheet{reference: ref}, sheet_name, column) do
    case UmyaNative.get_column_width(ref, sheet_name, column) do
      width when is_float(width) -> {:ok, width}
      {:error, reason} -> {:error, reason}
      result -> result
    end
  end

  @doc """
  Gets whether a column width is automatically adjusted.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `column` - The column letter (e.g., "A", "B")

  ## Returns

  - `{:ok, auto_width}` on success where auto_width is a boolean
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, auto_width} = UmyaSpreadsheet.RowColumnFunctions.get_column_auto_width(spreadsheet, "Sheet1", "A")
  """
  def get_column_auto_width(%Spreadsheet{reference: ref}, sheet_name, column) do
    case UmyaNative.get_column_auto_width(ref, sheet_name, column) do
      auto_width when is_boolean(auto_width) -> {:ok, auto_width}
      {:error, reason} -> {:error, reason}
      result -> result
    end
  end

  @doc """
  Gets whether a column is hidden.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `column` - The column letter (e.g., "A", "B")

  ## Returns

  - `{:ok, hidden}` on success where hidden is a boolean
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, hidden} = UmyaSpreadsheet.RowColumnFunctions.get_column_hidden(spreadsheet, "Sheet1", "A")
  """
  def get_column_hidden(%Spreadsheet{reference: ref}, sheet_name, column) do
    case UmyaNative.get_column_hidden(ref, sheet_name, column) do
      hidden when is_boolean(hidden) -> {:ok, hidden}
      {:error, reason} -> {:error, reason}
      result -> result
    end
  end

  @doc """
  Gets the height of a row.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `row_number` - The row number (1-based)

  ## Returns

  - `{:ok, height}` on success where height is a float
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, height} = UmyaSpreadsheet.RowColumnFunctions.get_row_height(spreadsheet, "Sheet1", 1)
  """
  def get_row_height(%Spreadsheet{reference: ref}, sheet_name, row_number) do
    case UmyaNative.get_row_height(ref, sheet_name, row_number) do
      height when is_float(height) -> {:ok, height}
      {:error, reason} -> {:error, reason}
      result -> result
    end
  end

  @doc """
  Gets whether a row is hidden.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `row_number` - The row number (1-based)

  ## Returns

  - `{:ok, hidden}` on success where hidden is a boolean
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, hidden} = UmyaSpreadsheet.RowColumnFunctions.get_row_hidden(spreadsheet, "Sheet1", 1)
  """
  def get_row_hidden(%Spreadsheet{reference: ref}, sheet_name, row_number) do
    case UmyaNative.get_row_hidden(ref, sheet_name, row_number) do
      hidden when is_boolean(hidden) -> {:ok, hidden}
      {:error, reason} -> {:error, reason}
      result -> result
    end
  end
end
