defmodule UmyaSpreadsheet.SheetFunctions do
  @moduledoc """
  Functions for manipulating sheets in a spreadsheet.
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaNative

  @doc """
  Gets a list of all sheet names in the spreadsheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct

  ## Returns

  - A list of sheet names

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      ["Sheet1", "Sheet2"] = UmyaSpreadsheet.SheetFunctions.get_sheet_names(spreadsheet)
  """
  def get_sheet_names(%Spreadsheet{reference: ref}) do
    UmyaNative.get_sheet_names(ref)
  end

  @doc """
  Adds a new sheet to the spreadsheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name for the new sheet

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.SheetFunctions.add_sheet(spreadsheet, "NewSheet")
  """
  def add_sheet(%Spreadsheet{reference: ref}, sheet_name) do
    case UmyaNative.add_sheet(ref, sheet_name) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Clones an existing sheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `source_sheet_name` - The name of the sheet to clone
  - `new_sheet_name` - The name for the new sheet

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.SheetFunctions.clone_sheet(spreadsheet, "Sheet1", "Sheet1 Copy")
  """
  def clone_sheet(%Spreadsheet{reference: ref}, source_sheet_name, new_sheet_name) do
    case UmyaNative.clone_sheet(ref, source_sheet_name, new_sheet_name) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Removes a sheet from the spreadsheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet to remove

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.SheetFunctions.remove_sheet(spreadsheet, "Sheet3")
  """
  def remove_sheet(%Spreadsheet{reference: ref}, sheet_name) do
    case UmyaNative.remove_sheet(ref, sheet_name) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets the visibility state of a sheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `state` - The visibility state ("visible", "hidden", "veryhidden")

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.SheetFunctions.set_sheet_state(spreadsheet, "Sheet1", "hidden")
  """
  def set_sheet_state(%Spreadsheet{reference: ref}, sheet_name, state) do
    case UmyaNative.set_sheet_state(ref, sheet_name, state) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets the protection for a sheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `password` - Optional password for the protection
  - `is_protected` - Boolean indicating whether the sheet should be protected

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      # Protect with password
      :ok = UmyaSpreadsheet.SheetFunctions.set_sheet_protection(spreadsheet, "Sheet1", "password", true)
      # Remove protection
      :ok = UmyaSpreadsheet.SheetFunctions.set_sheet_protection(spreadsheet, "Sheet1", nil, false)
  """
  def set_sheet_protection(%Spreadsheet{reference: ref}, sheet_name, password, is_protected) do
    case UmyaNative.set_sheet_protection(ref, sheet_name, password, is_protected) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Moves a range of cells to a new position.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `range` - The range to move (e.g., "A1:B5")
  - `rows` - Number of rows to move (positive for down, negative for up)
  - `columns` - Number of columns to move (positive for right, negative for left)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      # Move range down 2 rows and right 3 columns
      :ok = UmyaSpreadsheet.SheetFunctions.move_range(spreadsheet, "Sheet1", "A1:B5", 2, 3)
  """
  def move_range(%Spreadsheet{reference: ref}, sheet_name, range, rows, columns) do
    case UmyaNative.move_range(ref, sheet_name, range, rows, columns) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Merges cells in a specified range.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `range` - The range to merge (e.g., "A1:B5")

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.SheetFunctions.add_merge_cells(spreadsheet, "Sheet1", "A1:B2")
  """
  def add_merge_cells(%Spreadsheet{reference: ref}, sheet_name, range) do
    case UmyaNative.add_merge_cells(ref, sheet_name, range) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Inserts new rows into a sheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `row_index` - The index where rows should be inserted
  - `amount` - The number of rows to insert

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      # Insert 2 new rows at row 3
      :ok = UmyaSpreadsheet.SheetFunctions.insert_new_row(spreadsheet, "Sheet1", 3, 2)
  """
  def insert_new_row(%Spreadsheet{reference: ref}, sheet_name, row_index, amount) do
    case UmyaNative.insert_new_row(ref, sheet_name, row_index, amount) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Inserts new columns into a sheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `column` - The column letter where columns should be inserted (e.g., "C")
  - `amount` - The number of columns to insert

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      # Insert 2 new columns at column C
      :ok = UmyaSpreadsheet.SheetFunctions.insert_new_column(spreadsheet, "Sheet1", "C", 2)
  """
  def insert_new_column(%Spreadsheet{reference: ref}, sheet_name, column, amount) do
    case UmyaNative.insert_new_column(ref, sheet_name, column, amount) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Inserts new columns into a sheet using column index.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `column_index` - The column index (1-based) where columns should be inserted
  - `amount` - The number of columns to insert

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      # Insert 2 new columns at column index 3 (column C)
      :ok = UmyaSpreadsheet.SheetFunctions.insert_new_column_by_index(spreadsheet, "Sheet1", 3, 2)
  """
  def insert_new_column_by_index(%Spreadsheet{reference: ref}, sheet_name, column_index, amount) do
    case UmyaNative.insert_new_column_by_index(ref, sheet_name, column_index, amount) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Removes rows from a sheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `row_index` - The index where rows should be removed from
  - `amount` - The number of rows to remove

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      # Remove 2 rows starting at row 3
      :ok = UmyaSpreadsheet.SheetFunctions.remove_row(spreadsheet, "Sheet1", 3, 2)
  """
  def remove_row(%Spreadsheet{reference: ref}, sheet_name, row_index, amount) do
    case UmyaNative.remove_row(ref, sheet_name, row_index, amount) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Removes columns from a sheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `column` - The column letter where columns should be removed from (e.g., "C")
  - `amount` - The number of columns to remove

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      # Remove 2 columns starting at column C
      :ok = UmyaSpreadsheet.SheetFunctions.remove_column(spreadsheet, "Sheet1", "C", 2)
  """
  def remove_column(%Spreadsheet{reference: ref}, sheet_name, column, amount) do
    case UmyaNative.remove_column(ref, sheet_name, column, amount) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Removes columns from a sheet using column index.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `column_index` - The column index (1-based) where columns should be removed from
  - `amount` - The number of columns to remove

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      # Remove 2 columns starting at column index 3 (column C)
      :ok = UmyaSpreadsheet.SheetFunctions.remove_column_by_index(spreadsheet, "Sheet1", 3, 2)
  """
  def remove_column_by_index(%Spreadsheet{reference: ref}, sheet_name, column_index, amount) do
    case UmyaNative.remove_column_by_index(ref, sheet_name, column_index, amount) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end
end
