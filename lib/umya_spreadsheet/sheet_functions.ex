defmodule UmyaSpreadsheet.SheetFunctions do
  @moduledoc """
  Functions for manipulating sheets in a spreadsheet.
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaSpreadsheet.ErrorHandling
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
    UmyaNative.add_sheet(ref, sheet_name)
    |> ErrorHandling.standardize_result()
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
    UmyaNative.clone_sheet(ref, source_sheet_name, new_sheet_name)
    |> ErrorHandling.standardize_result()
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
    UmyaNative.remove_sheet(ref, sheet_name)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Renames an existing sheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `old_sheet_name` - The current name of the sheet
  - `new_sheet_name` - The new name for the sheet

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.SheetFunctions.rename_sheet(spreadsheet, "Sheet1", "Updated Sheet")
  """
  def rename_sheet(%Spreadsheet{reference: ref}, old_sheet_name, new_sheet_name) do
    UmyaNative.rename_sheet(ref, old_sheet_name, new_sheet_name)
    |> ErrorHandling.standardize_result()
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
    UmyaNative.set_sheet_state(ref, sheet_name, state)
    |> ErrorHandling.standardize_result()
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
    UmyaNative.set_sheet_protection(ref, sheet_name, password, is_protected)
    |> ErrorHandling.standardize_result()
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
    UmyaNative.move_range(ref, sheet_name, range, rows, columns)
    |> ErrorHandling.standardize_result()
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
    UmyaNative.add_merge_cells(ref, sheet_name, range)
    |> ErrorHandling.standardize_result()
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
    UmyaNative.insert_new_row(ref, sheet_name, row_index, amount)
    |> ErrorHandling.standardize_result()
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
    UmyaNative.insert_new_column(ref, sheet_name, column, amount)
    |> ErrorHandling.standardize_result()
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
    UmyaNative.insert_new_column_by_index(ref, sheet_name, column_index, amount)
    |> ErrorHandling.standardize_result()
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
      # Remove 2 rows starting from row 3
      :ok = UmyaSpreadsheet.SheetFunctions.remove_row(spreadsheet, "Sheet1", 3, 2)
  """
  def remove_row(%Spreadsheet{reference: ref}, sheet_name, row_index, amount) do
    UmyaNative.remove_row(ref, sheet_name, row_index, amount)
    |> ErrorHandling.standardize_result()
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
      # Remove 2 columns starting from column C
      :ok = UmyaSpreadsheet.SheetFunctions.remove_column(spreadsheet, "Sheet1", "C", 2)
  """
  def remove_column(%Spreadsheet{reference: ref}, sheet_name, column, amount) do
    UmyaNative.remove_column(ref, sheet_name, column, amount)
    |> ErrorHandling.standardize_result()
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
      # Remove 2 columns starting from column index 3 (column C)
      :ok = UmyaSpreadsheet.SheetFunctions.remove_column_by_index(spreadsheet, "Sheet1", 3, 2)
  """
  def remove_column_by_index(%Spreadsheet{reference: ref}, sheet_name, column_index, amount) do
    UmyaNative.remove_column_by_index(ref, sheet_name, column_index, amount)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the total number of sheets in the spreadsheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct

  ## Returns

  - The number of sheets as an integer

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      3 = UmyaSpreadsheet.SheetFunctions.get_sheet_count(spreadsheet)
  """
  def get_sheet_count(%Spreadsheet{reference: ref}) do
    UmyaNative.get_sheet_count(ref)
  end

  @doc """
  Gets the currently active sheet tab index.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct

  ## Returns

  - `{:ok, index}` on success where index is the 0-based active tab index
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, 0} = UmyaSpreadsheet.SheetFunctions.get_active_sheet(spreadsheet)
  """
  def get_active_sheet(%Spreadsheet{reference: ref}) do
    UmyaNative.get_active_tab(ref)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the visibility state of a sheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet

  ## Returns

  - `{:ok, state}` on success where state is "visible", "hidden", or "veryhidden"
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, "visible"} = UmyaSpreadsheet.SheetFunctions.get_sheet_state(spreadsheet, "Sheet1")
  """
  def get_sheet_state(%Spreadsheet{reference: ref}, sheet_name) do
    # Note: This would need a native implementation if not already available
    # For now, we'll mark this as a placeholder that needs native implementation
    case UmyaNative.get_sheet_state(ref, sheet_name) do
      {:ok, state} -> {:ok, state}
      error -> ErrorHandling.standardize_result(error)
    end
  rescue
    UndefinedFunctionError -> {:error, "get_sheet_state native function not yet implemented"}
  end

  @doc """
  Checks if a sheet is protected and gets protection details.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet

  ## Returns

  - `{:ok, %{protected: boolean, details: map}}` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, %{protected: true, details: %{}}} = UmyaSpreadsheet.SheetFunctions.get_sheet_protection(spreadsheet, "Sheet1")
  """
  def get_sheet_protection(%Spreadsheet{reference: ref}, sheet_name) do
    # Note: This would need a native implementation if not already available
    # For now, we'll mark this as a placeholder that needs native implementation
    case UmyaNative.get_sheet_protection(ref, sheet_name) do
      {:ok, details} -> {:ok, details}
      error -> ErrorHandling.standardize_result(error)
    end
  rescue
    UndefinedFunctionError -> {:error, "get_sheet_protection native function not yet implemented"}
  end

  @doc """
  Gets the list of merged cell ranges in a sheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet

  ## Returns

  - `{:ok, ranges}` on success where ranges is a list of range strings
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, ["A1:B2", "C3:D4"]} = UmyaSpreadsheet.SheetFunctions.get_merge_cells(spreadsheet, "Sheet1")
  """
  def get_merge_cells(%Spreadsheet{reference: ref}, sheet_name) do
    # Note: This would need a native implementation if not already available
    # For now, we'll mark this as a placeholder that needs native implementation
    case UmyaNative.get_merge_cells(ref, sheet_name) do
      {:ok, ranges} -> {:ok, ranges}
      error -> ErrorHandling.standardize_result(error)
    end
  rescue
    UndefinedFunctionError -> {:error, "get_merge_cells native function not yet implemented"}
  end
end
