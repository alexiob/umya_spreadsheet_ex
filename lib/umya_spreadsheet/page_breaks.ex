defmodule UmyaSpreadsheet.PageBreaks do
  @moduledoc """
  Functions for managing page breaks in worksheets.

  Page breaks control where pages split when printing or viewing in Page Break Preview.
  Excel supports both manual (user-defined) and automatic page breaks.

  ## Manual vs Automatic Page Breaks

  - **Manual page breaks**: Explicitly set by the user and remain fixed
  - **Automatic page breaks**: Calculated by Excel based on page size and content

  Manual page breaks have priority over automatic ones and are useful for:
  - Ensuring specific sections appear on separate pages
  - Creating consistent report layouts
  - Controlling print output formatting

  ## Usage

  Page breaks are identified by row or column numbers:
  - Row page breaks: Insert a horizontal page break above the specified row
  - Column page breaks: Insert a vertical page break to the left of the specified column

  All functions require a valid sheet name and will return `{:error, :sheet_not_found}`
  if the sheet doesn't exist.
  """

  alias UmyaSpreadsheet.Spreadsheet

  @doc """
  Adds a manual row page break at the specified row number.

  A row page break creates a horizontal line where pages split. The break is inserted
  above the specified row, meaning content from that row onwards will appear on a new page.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `row_number` - The row number where the page break should be inserted (1-based)
  - `manual` - Whether this is a manual page break (default: true)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      # Add a manual page break above row 25
      :ok = UmyaSpreadsheet.PageBreaks.add_row_page_break(spreadsheet, "Sheet1", 25)

      # Add an automatic page break above row 50
      :ok = UmyaSpreadsheet.PageBreaks.add_row_page_break(spreadsheet, "Sheet1", 50, false)

  """
  @spec add_row_page_break(Spreadsheet.t(), String.t(), pos_integer(), boolean()) ::
          :ok | {:error, atom()}
  def add_row_page_break(%Spreadsheet{reference: ref}, sheet_name, row_number, manual \\ true) do
    case UmyaNative.add_row_page_break(ref, sheet_name, row_number, manual) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Adds a manual column page break at the specified column number.

  A column page break creates a vertical line where pages split. The break is inserted
  to the left of the specified column, meaning content from that column onwards will
  appear on a new page.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `column_number` - The column number where the page break should be inserted (1-based)
  - `manual` - Whether this is a manual page break (default: true)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      # Add a manual page break to the left of column 10 (column J)
      :ok = UmyaSpreadsheet.PageBreaks.add_column_page_break(spreadsheet, "Sheet1", 10)

      # Add an automatic page break to the left of column 15 (column O)
      :ok = UmyaSpreadsheet.PageBreaks.add_column_page_break(spreadsheet, "Sheet1", 15, false)

  """
  @spec add_column_page_break(Spreadsheet.t(), String.t(), pos_integer(), boolean()) ::
          :ok | {:error, atom()}
  def add_column_page_break(
        %Spreadsheet{reference: ref},
        sheet_name,
        column_number,
        manual \\ true
      ) do
    case UmyaNative.add_column_page_break(ref, sheet_name, column_number, manual) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Removes a row page break at the specified row number.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `row_number` - The row number where the page break should be removed (1-based)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      # Remove page break above row 25
      :ok = UmyaSpreadsheet.PageBreaks.remove_row_page_break(spreadsheet, "Sheet1", 25)

  """
  @spec remove_row_page_break(Spreadsheet.t(), String.t(), pos_integer()) ::
          :ok | {:error, atom()}
  def remove_row_page_break(%Spreadsheet{reference: ref}, sheet_name, row_number) do
    case UmyaNative.remove_row_page_break(ref, sheet_name, row_number) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Removes a column page break at the specified column number.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `column_number` - The column number where the page break should be removed (1-based)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      # Remove page break to the left of column 10
      :ok = UmyaSpreadsheet.PageBreaks.remove_column_page_break(spreadsheet, "Sheet1", 10)

  """
  @spec remove_column_page_break(Spreadsheet.t(), String.t(), pos_integer()) ::
          :ok | {:error, atom()}
  def remove_column_page_break(%Spreadsheet{reference: ref}, sheet_name, column_number) do
    case UmyaNative.remove_column_page_break(ref, sheet_name, column_number) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Gets all row page breaks for the specified sheet.

  Returns a list of tuples containing the row number and whether it's a manual page break.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet

  ## Returns

  - `{:ok, breaks}` on success, where breaks is a list of `{row_number, is_manual}` tuples
  - `{:error, reason}` on failure

  ## Examples

      # Get all row page breaks
      {:ok, breaks} = UmyaSpreadsheet.PageBreaks.get_row_page_breaks(spreadsheet, "Sheet1")
      # breaks might be: [{25, true}, {50, false}, {75, true}]

      # Check if there are any manual breaks
      manual_breaks = Enum.filter(breaks, fn {_row, manual} -> manual end)

  """
  @spec get_row_page_breaks(Spreadsheet.t(), String.t()) ::
          {:ok, [{pos_integer(), boolean()}]} | {:error, atom()}
  def get_row_page_breaks(%Spreadsheet{reference: ref}, sheet_name) do
    case UmyaNative.get_row_page_breaks(ref, sheet_name) do
      result when is_list(result) -> {:ok, result}
      result -> result
    end
  end

  @doc """
  Gets all column page breaks for the specified sheet.

  Returns a list of tuples containing the column number and whether it's a manual page break.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet

  ## Returns

  - `{:ok, breaks}` on success, where breaks is a list of `{column_number, is_manual}` tuples
  - `{:error, reason}` on failure

  ## Examples

      # Get all column page breaks
      {:ok, breaks} = UmyaSpreadsheet.PageBreaks.get_column_page_breaks(spreadsheet, "Sheet1")
      # breaks might be: [{10, true}, {20, false}]

      # Get only manual column breaks
      manual_breaks = Enum.filter(breaks, fn {_col, manual} -> manual end)

  """
  @spec get_column_page_breaks(Spreadsheet.t(), String.t()) ::
          {:ok, [{pos_integer(), boolean()}]} | {:error, atom()}
  def get_column_page_breaks(%Spreadsheet{reference: ref}, sheet_name) do
    case UmyaNative.get_column_page_breaks(ref, sheet_name) do
      result when is_list(result) -> {:ok, result}
      result -> result
    end
  end

  @doc """
  Clears all row page breaks for the specified sheet.

  This removes both manual and automatic row page breaks.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      # Clear all row page breaks
      :ok = UmyaSpreadsheet.PageBreaks.clear_row_page_breaks(spreadsheet, "Sheet1")

  """
  @spec clear_row_page_breaks(Spreadsheet.t(), String.t()) :: :ok | {:error, atom()}
  def clear_row_page_breaks(%Spreadsheet{reference: ref}, sheet_name) do
    case UmyaNative.clear_row_page_breaks(ref, sheet_name) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Clears all column page breaks for the specified sheet.

  This removes both manual and automatic column page breaks.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      # Clear all column page breaks
      :ok = UmyaSpreadsheet.PageBreaks.clear_column_page_breaks(spreadsheet, "Sheet1")

  """
  @spec clear_column_page_breaks(Spreadsheet.t(), String.t()) :: :ok | {:error, atom()}
  def clear_column_page_breaks(%Spreadsheet{reference: ref}, sheet_name) do
    case UmyaNative.clear_column_page_breaks(ref, sheet_name) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Checks if a row page break exists at the specified row number.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `row_number` - The row number to check (1-based)

  ## Returns

  - `{:ok, true}` if a page break exists at the row
  - `{:ok, false}` if no page break exists at the row
  - `{:error, reason}` on failure

  ## Examples

      # Check if there's a page break above row 25
      {:ok, has_break} = UmyaSpreadsheet.PageBreaks.has_row_page_break(spreadsheet, "Sheet1", 25)

      if has_break do
        IO.puts("Page break found at row 25")
      end

  """
  @spec has_row_page_break(Spreadsheet.t(), String.t(), pos_integer()) ::
          {:ok, boolean()} | {:error, atom()}
  def has_row_page_break(%Spreadsheet{reference: ref}, sheet_name, row_number) do
    case UmyaNative.has_row_page_break(ref, sheet_name, row_number) do
      result when is_boolean(result) -> {:ok, result}
      result -> result
    end
  end

  @doc """
  Checks if a column page break exists at the specified column number.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `column_number` - The column number to check (1-based)

  ## Returns

  - `{:ok, true}` if a page break exists at the column
  - `{:ok, false}` if no page break exists at the column
  - `{:error, reason}` on failure

  ## Examples

      # Check if there's a page break to the left of column 10
      {:ok, has_break} = UmyaSpreadsheet.PageBreaks.has_column_page_break(spreadsheet, "Sheet1", 10)

      if has_break do
        IO.puts("Page break found at column 10")
      end

  """
  @spec has_column_page_break(Spreadsheet.t(), String.t(), pos_integer()) ::
          {:ok, boolean()} | {:error, atom()}
  def has_column_page_break(%Spreadsheet{reference: ref}, sheet_name, column_number) do
    case UmyaNative.has_column_page_break(ref, sheet_name, column_number) do
      result when is_boolean(result) -> {:ok, result}
      result -> result
    end
  end

  @doc """
  Gets the count of row page breaks for the specified sheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet

  ## Returns

  - `{:ok, count}` on success, where count is the number of row page breaks
  - `{:error, reason}` on failure

  ## Examples

      {:ok, count} = UmyaSpreadsheet.PageBreaks.count_row_page_breaks(spreadsheet, "Sheet1")
      IO.puts("Total row page breaks: \#{count}")

  """
  @spec count_row_page_breaks(Spreadsheet.t(), String.t()) ::
          {:ok, non_neg_integer()} | {:error, atom()}
  def count_row_page_breaks(spreadsheet, sheet_name) do
    case get_row_page_breaks(spreadsheet, sheet_name) do
      {:ok, breaks} -> {:ok, length(breaks)}
      error -> error
    end
  end

  @doc """
  Gets the count of column page breaks for the specified sheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet

  ## Returns

  - `{:ok, count}` on success, where count is the number of column page breaks
  - `{:error, reason}` on failure

  ## Examples

      {:ok, count} = UmyaSpreadsheet.PageBreaks.count_column_page_breaks(spreadsheet, "Sheet1")
      IO.puts("Total column page breaks: \#{count}")

  """
  @spec count_column_page_breaks(Spreadsheet.t(), String.t()) ::
          {:ok, non_neg_integer()} | {:error, atom()}
  def count_column_page_breaks(spreadsheet, sheet_name) do
    case get_column_page_breaks(spreadsheet, sheet_name) do
      {:ok, breaks} -> {:ok, length(breaks)}
      error -> error
    end
  end

  # Bulk Operations & Convenience Functions

  @doc """
  Adds multiple row page breaks in a single operation for better performance.

  This function efficiently adds multiple row page breaks without needing separate
  calls for each break. All breaks are added as manual page breaks.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `row_numbers` - A list of row numbers where page breaks should be inserted (1-based)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      # Add page breaks above rows 10, 20, and 30
      :ok = UmyaSpreadsheet.PageBreaks.add_row_page_breaks(spreadsheet, "Sheet1", [10, 20, 30])

      # Add page breaks for section headers
      section_rows = [15, 45, 75, 105]
      :ok = UmyaSpreadsheet.PageBreaks.add_row_page_breaks(spreadsheet, "Reports", section_rows)

  """
  @spec add_row_page_breaks(Spreadsheet.t(), String.t(), [pos_integer()]) ::
          :ok | {:error, atom()}
  def add_row_page_breaks(spreadsheet, sheet_name, row_numbers) when is_list(row_numbers) do
    Enum.reduce_while(row_numbers, :ok, fn row_number, _acc ->
      case add_row_page_break(spreadsheet, sheet_name, row_number) do
        :ok -> {:cont, :ok}
        error -> {:halt, error}
      end
    end)
  end

  @doc """
  Adds multiple column page breaks in a single operation for better performance.

  This function efficiently adds multiple column page breaks without needing separate
  calls for each break. All breaks are added as manual page breaks.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `column_numbers` - A list of column numbers where page breaks should be inserted (1-based)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      # Add page breaks before columns 5, 10, and 15
      :ok = UmyaSpreadsheet.PageBreaks.add_column_page_breaks(spreadsheet, "Sheet1", [5, 10, 15])

      # Add page breaks for different data sections
      section_columns = [6, 12, 18, 24]
      :ok = UmyaSpreadsheet.PageBreaks.add_column_page_breaks(spreadsheet, "Data", section_columns)

  """
  @spec add_column_page_breaks(Spreadsheet.t(), String.t(), [pos_integer()]) ::
          :ok | {:error, atom()}
  def add_column_page_breaks(spreadsheet, sheet_name, column_numbers)
      when is_list(column_numbers) do
    Enum.reduce_while(column_numbers, :ok, fn column_number, _acc ->
      case add_column_page_break(spreadsheet, sheet_name, column_number) do
        :ok -> {:cont, :ok}
        error -> {:halt, error}
      end
    end)
  end

  @doc """
  Removes multiple row page breaks efficiently in a single operation.

  This function efficiently removes multiple row page breaks without needing separate
  calls for each break removal.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `row_numbers` - A list of row numbers where page breaks should be removed (1-based)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      # Remove page breaks from specific rows
      :ok = UmyaSpreadsheet.PageBreaks.remove_row_page_breaks(spreadsheet, "Sheet1", [10, 20, 30])

      # Remove page breaks from old section positions
      old_sections = [25, 50, 75]
      :ok = UmyaSpreadsheet.PageBreaks.remove_row_page_breaks(spreadsheet, "Reports", old_sections)

  """
  @spec remove_row_page_breaks(Spreadsheet.t(), String.t(), [pos_integer()]) ::
          :ok | {:error, atom()}
  def remove_row_page_breaks(spreadsheet, sheet_name, row_numbers) when is_list(row_numbers) do
    Enum.reduce_while(row_numbers, :ok, fn row_number, _acc ->
      case remove_row_page_break(spreadsheet, sheet_name, row_number) do
        :ok -> {:cont, :ok}
        error -> {:halt, error}
      end
    end)
  end

  @doc """
  Removes multiple column page breaks efficiently in a single operation.

  This function efficiently removes multiple column page breaks without needing separate
  calls for each break removal.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `column_numbers` - A list of column numbers where page breaks should be removed (1-based)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      # Remove page breaks from specific columns
      :ok = UmyaSpreadsheet.PageBreaks.remove_column_page_breaks(spreadsheet, "Sheet1", [5, 10, 15])

      # Remove page breaks from old column positions
      old_columns = [8, 16, 24]
      :ok = UmyaSpreadsheet.PageBreaks.remove_column_page_breaks(spreadsheet, "Data", old_columns)

  """
  @spec remove_column_page_breaks(Spreadsheet.t(), String.t(), [pos_integer()]) ::
          :ok | {:error, atom()}
  def remove_column_page_breaks(spreadsheet, sheet_name, column_numbers)
      when is_list(column_numbers) do
    Enum.reduce_while(column_numbers, :ok, fn column_number, _acc ->
      case remove_column_page_break(spreadsheet, sheet_name, column_number) do
        :ok -> {:cont, :ok}
        error -> {:halt, error}
      end
    end)
  end

  @doc """
  Clears all page breaks (both row and column) from the specified sheet.

  This is a convenience function that removes all manual page breaks from a worksheet,
  effectively resetting the page break layout to automatic only.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      # Clear all page breaks from a sheet
      :ok = UmyaSpreadsheet.PageBreaks.clear_all_page_breaks(spreadsheet, "Sheet1")

      # Reset page breaks when reformatting a report
      :ok = UmyaSpreadsheet.PageBreaks.clear_all_page_breaks(spreadsheet, "Monthly Report")

  """
  @spec clear_all_page_breaks(Spreadsheet.t(), String.t()) :: :ok | {:error, atom()}
  def clear_all_page_breaks(spreadsheet, sheet_name) do
    with :ok <- clear_row_page_breaks(spreadsheet, sheet_name),
         :ok <- clear_column_page_breaks(spreadsheet, sheet_name) do
      :ok
    end
  end

  @doc """
  Gets all page breaks (both row and column) for the specified sheet in a single result.

  This convenience function retrieves both row and column page breaks, returning them
  in a structured format for easy access to all page break information.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet

  ## Returns

  - `{:ok, %{row_breaks: [row_numbers], column_breaks: [column_numbers]}}` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, %{row_breaks: rows, column_breaks: cols}} =
        UmyaSpreadsheet.PageBreaks.get_all_page_breaks(spreadsheet, "Sheet1")

      IO.puts("Row breaks at: \#{inspect(rows)}")
      IO.puts("Column breaks at: \#{inspect(cols)}")

      # Check if any page breaks exist
      {:ok, breaks} = UmyaSpreadsheet.PageBreaks.get_all_page_breaks(spreadsheet, "Report")
      has_breaks = length(breaks.row_breaks) > 0 or length(breaks.column_breaks) > 0

  """
  @spec get_all_page_breaks(Spreadsheet.t(), String.t()) ::
          {:ok, %{row_breaks: [pos_integer()], column_breaks: [pos_integer()]}} | {:error, atom()}
  def get_all_page_breaks(spreadsheet, sheet_name) do
    with {:ok, row_breaks} <- get_row_page_breaks(spreadsheet, sheet_name),
         {:ok, column_breaks} <- get_column_page_breaks(spreadsheet, sheet_name) do
      {:ok, %{row_breaks: row_breaks, column_breaks: column_breaks}}
    end
  end
end
