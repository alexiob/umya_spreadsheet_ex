defmodule UmyaSpreadsheet.Table do
  @moduledoc """
  Functions for creating and managing Excel tables.

  Excel tables are structured data ranges that provide built-in filtering, sorting,
  and formatting capabilities. This module provides functions to:

  * Create tables from data ranges
  * Configure table columns and styling
  * Manage table filters and totals rows
  * Modify existing table structures
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaNative

  @doc """
  Adds a new table to a spreadsheet.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet where the table will be created
    * `table_name` - Internal name of the table (must be unique within the workbook)
    * `display_name` - Display name shown to users
    * `start_cell` - Top-left cell of the table range (e.g., "A1")
    * `end_cell` - Bottom-right cell of the table range (e.g., "D10")
    * `columns` - List of column names for the table headers
    * `has_totals_row` - Optional boolean indicating if the table should include a totals row

  ## Examples

  ```elixir
  # Create a sales data table
  {:ok, :ok} = Table.add_table(
    spreadsheet,
    "Sheet1",
    "SalesTable",
    "Sales Data",
    "A1",
    "D10",
    ["Region", "Product", "Sales", "Date"],
    true  # Include totals row
  )
  ```

  Returns `{:ok, :ok}` on success or `{:error, reason}` on failure.
  """
  @spec add_table(
          Spreadsheet.t(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          [String.t()],
          boolean() | nil
        ) :: :ok | {:error, String.t()}
  def add_table(
        %Spreadsheet{reference: resource},
        sheet_name,
        table_name,
        display_name,
        start_cell,
        end_cell,
        columns,
        has_totals_row \\ nil
      ) do
    UmyaNative.add_table(
      resource,
      sheet_name,
      table_name,
      display_name,
      start_cell,
      end_cell,
      columns,
      has_totals_row
    )
  end

  @doc """
  Gets all tables from a worksheet.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet to get tables from

  ## Returns

  A list of maps, each containing table information with keys:
  - `"name"` - Internal table name
  - `"display_name"` - Display name
  - `"start_cell"` - Top-left cell reference
  - `"end_cell"` - Bottom-right cell reference
  - `"columns"` - List of column names
  - `"has_totals_row"` - Boolean indicating if totals row is enabled
  - `"style_info"` - Map with styling information (if table has custom styling)

  ## Examples

  ```elixir
  {:ok, tables} = Table.get_tables(spreadsheet, "Sheet1")
  # tables = [%{"name" => "SalesTable", "display_name" => "Sales Data", ...}]
  ```
  """
  @spec get_tables(Spreadsheet.t(), String.t()) :: {:ok, [map()]} | {:error, String.t()}
  def get_tables(%Spreadsheet{reference: resource}, sheet_name) do
    UmyaNative.get_tables(resource, sheet_name)
  end

  @doc """
  Removes a table from a worksheet by name.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet containing the table
    * `table_name` - Internal name of the table to remove

  ## Examples

  ```elixir
  Table.remove_table(spreadsheet, "Sheet1", "SalesTable")
  ```

  Returns `:ok` on success or `{:error, reason}` on failure.
  """
  @spec remove_table(Spreadsheet.t(), String.t(), String.t()) :: {:ok, :ok} | {:error, String.t()}
  def remove_table(%Spreadsheet{reference: resource}, sheet_name, table_name) do
    UmyaNative.remove_table(resource, sheet_name, table_name)
  end

  @doc """
  Checks if a worksheet has any tables.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet to check

  ## Examples

  ```elixir
  {:ok, has_tables?} = Table.has_tables(spreadsheet, "Sheet1")
  # has_tables? = true or false
  ```

  Returns `{:ok, boolean}` if successful, `{:error, reason}` otherwise.
  """
  @spec has_tables(Spreadsheet.t(), String.t()) :: {:ok, boolean()} | {:error, String.t()}
  def has_tables(%Spreadsheet{reference: resource}, sheet_name) do
    UmyaNative.has_tables(resource, sheet_name)
  end

  @doc """
  Counts the number of tables in a worksheet.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet to count tables in

  ## Examples

  ```elixir
  {:ok, count} = Table.count_tables(spreadsheet, "Sheet1")
  # count = 3
  ```

  Returns `{:ok, integer}` with the number of tables on success.
  """
  @spec count_tables(Spreadsheet.t(), String.t()) ::
          {:ok, non_neg_integer()} | {:error, String.t()}
  def count_tables(%Spreadsheet{reference: resource}, sheet_name) do
    UmyaNative.count_tables(resource, sheet_name)
  end

  @doc """
  Sets table style information.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet containing the table
    * `table_name` - Internal name of the table to style
    * `style_name` - Name of the table style to apply
    * `show_first_col` - Boolean to highlight the first column
    * `show_last_col` - Boolean to highlight the last column
    * `show_row_stripes` - Boolean to show alternating row colors
    * `show_col_stripes` - Boolean to show alternating column colors

  ## Examples

  ```elixir
  Table.set_table_style(
    spreadsheet,
    "Sheet1",
    "SalesTable",
    "TableStyleLight1",
    true,   # Highlight first column
    true,   # Highlight last column
    true,   # Show row stripes
    false   # Don't show column stripes
  )
  ```

  Returns `:ok` on success or `{:error, reason}` on failure.
  """
  @spec set_table_style(
          Spreadsheet.t(),
          String.t(),
          String.t(),
          String.t(),
          boolean(),
          boolean(),
          boolean(),
          boolean()
        ) :: {:ok, :ok} | {:error, String.t()}
  def set_table_style(
        %Spreadsheet{reference: resource},
        sheet_name,
        table_name,
        style_name,
        show_first_col,
        show_last_col,
        show_row_stripes,
        show_col_stripes
      ) do
    UmyaNative.set_table_style(
      resource,
      sheet_name,
      table_name,
      style_name,
      show_first_col,
      show_last_col,
      show_row_stripes,
      show_col_stripes
    )
  end

  @doc """
  Removes table style information, reverting to default styling.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet containing the table
    * `table_name` - Internal name of the table to remove styling from

  ## Examples

  ```elixir
  Table.remove_table_style(spreadsheet, "Sheet1", "SalesTable")
  ```

  Returns `:ok` on success or `{:error, reason}` on failure.
  """
  @spec remove_table_style(Spreadsheet.t(), String.t(), String.t()) ::
          {:ok, :ok} | {:error, String.t()}
  def remove_table_style(%Spreadsheet{reference: resource}, sheet_name, table_name) do
    UmyaNative.remove_table_style(resource, sheet_name, table_name)
  end

  @doc """
  Adds a new column to an existing table.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet containing the table
    * `table_name` - Internal name of the table to modify
    * `column_name` - Name of the new column
    * `totals_row_function` - Optional function for the totals row ("sum", "count", "average", etc.)
    * `totals_row_label` - Optional custom label for the totals row

  ## Examples

  ```elixir
  Table.add_table_column(
    spreadsheet,
    "Sheet1",
    "SalesTable",
    "Total",
    "sum",
    "Grand Total"
  )
  ```

  Returns `:ok` on success or `{:error, reason}` on failure.
  """
  @spec add_table_column(
          Spreadsheet.t(),
          String.t(),
          String.t(),
          String.t(),
          String.t() | nil,
          String.t() | nil
        ) :: {:ok, :ok} | {:error, String.t()}
  def add_table_column(
        %Spreadsheet{reference: resource},
        sheet_name,
        table_name,
        column_name,
        totals_row_function \\ nil,
        totals_row_label \\ nil
      ) do
    UmyaNative.add_table_column(
      resource,
      sheet_name,
      table_name,
      column_name,
      totals_row_function,
      totals_row_label
    )
  end

  @doc """
  Modifies an existing table column.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet containing the table
    * `table_name` - Internal name of the table to modify
    * `old_column_name` - Current name of the column to modify
    * `new_column_name` - Optional new name for the column
    * `totals_row_function` - Optional function for the totals row
    * `totals_row_label` - Optional custom label for the totals row

  ## Examples

  ```elixir
  # Rename a column and change its totals function
  Table.modify_table_column(
    spreadsheet,
    "Sheet1",
    "SalesTable",
    "Sales",
    "Revenue",
    "sum",
    "Total Revenue"
  )
  ```

  Returns `:ok` on success or `{:error, reason}` on failure.
  """
  @spec modify_table_column(
          Spreadsheet.t(),
          String.t(),
          String.t(),
          String.t(),
          String.t() | nil,
          String.t() | nil,
          String.t() | nil
        ) :: {:ok, :ok} | {:error, String.t()}
  def modify_table_column(
        %Spreadsheet{reference: resource},
        sheet_name,
        table_name,
        old_column_name,
        new_column_name \\ nil,
        totals_row_function \\ nil,
        totals_row_label \\ nil
      ) do
    UmyaNative.modify_table_column(
      resource,
      sheet_name,
      table_name,
      old_column_name,
      new_column_name,
      totals_row_function,
      totals_row_label
    )
  end

  @doc """
  Sets the visibility of the totals row for a table.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet containing the table
    * `table_name` - Internal name of the table to modify
    * `show_totals_row` - Boolean indicating whether to show the totals row

  ## Examples

  ```elixir
  # Show the totals row
  Table.set_table_totals_row(spreadsheet, "Sheet1", "SalesTable", true)

  # Hide the totals row
  Table.set_table_totals_row(spreadsheet, "Sheet1", "SalesTable", false)
  ```

  Returns `:ok` on success or `{:error, reason}` on failure.
  """
  @spec set_table_totals_row(Spreadsheet.t(), String.t(), String.t(), boolean()) ::
          {:ok, :ok} | {:error, String.t()}
  def set_table_totals_row(
        %Spreadsheet{reference: resource},
        sheet_name,
        table_name,
        show_totals_row
      ) do
    UmyaNative.set_table_totals_row(resource, sheet_name, table_name, show_totals_row)
  end

  @doc """
  Gets a specific table by name from a worksheet.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet containing the table
    * `table_name` - Internal name of the table to retrieve

  ## Returns

  A map containing table information with keys:
  - `"name"` - Internal table name
  - `"display_name"` - Display name
  - `"start_cell"` - Top-left cell reference
  - `"end_cell"` - Bottom-right cell reference
  - `"columns"` - List of column names
  - `"has_totals_row"` - Boolean indicating if totals row is enabled
  - `"style_info"` - Map with styling information (if table has custom styling)

  ## Examples

  ```elixir
  {:ok, table} = Table.get_table(spreadsheet, "Sheet1", "SalesTable")
  # table = %{"name" => "SalesTable", "display_name" => "Sales Data", ...}
  ```
  """
  @spec get_table(Spreadsheet.t(), String.t(), String.t()) :: {:ok, map()} | {:error, String.t()}
  def get_table(%Spreadsheet{reference: resource}, sheet_name, table_name) do
    UmyaNative.get_table(resource, sheet_name, table_name)
  end

  @doc """
  Gets table style information for a specific table.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet containing the table
    * `table_name` - Internal name of the table

  ## Returns

  A map containing style information with keys:
  - `"name"` - Style name (e.g., "TableStyleLight1")
  - `"show_first_column"` - Boolean indicating if first column is highlighted
  - `"show_last_column"` - Boolean indicating if last column is highlighted
  - `"show_row_stripes"` - Boolean indicating if row stripes are shown
  - `"show_column_stripes"` - Boolean indicating if column stripes are shown

  ## Examples

  ```elixir
  {:ok, style} = Table.get_table_style(spreadsheet, "Sheet1", "SalesTable")
  # style = %{"name" => "TableStyleLight1", "show_first_column" => "true", ...}
  ```
  """
  @spec get_table_style(Spreadsheet.t(), String.t(), String.t()) ::
          {:ok, map()} | {:error, String.t()}
  def get_table_style(%Spreadsheet{reference: resource}, sheet_name, table_name) do
    UmyaNative.get_table_style(resource, sheet_name, table_name)
  end

  @doc """
  Gets column information for a specific table.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet containing the table
    * `table_name` - Internal name of the table

  ## Returns

  A list of maps, each containing column information with keys:
  - `"name"` - Column name
  - `"totals_row_label"` - Custom label for totals row (if set)
  - `"totals_row_function"` - Function used in totals row (e.g., "sum", "count")

  ## Examples

  ```elixir
  {:ok, columns} = Table.get_table_columns(spreadsheet, "Sheet1", "SalesTable")
  # columns = [%{"name" => "Region", "totals_row_function" => "none"}, ...]
  ```
  """
  @spec get_table_columns(Spreadsheet.t(), String.t(), String.t()) ::
          {:ok, [map()]} | {:error, String.t()}
  def get_table_columns(%Spreadsheet{reference: resource}, sheet_name, table_name) do
    UmyaNative.get_table_columns(resource, sheet_name, table_name)
  end

  @doc """
  Gets the totals row visibility status for a specific table.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet containing the table
    * `table_name` - Internal name of the table

  ## Examples

  ```elixir
  {:ok, has_totals_row?} = Table.get_table_totals_row(spreadsheet, "Sheet1", "SalesTable")
  # has_totals_row? = true or false
  ```

  Returns `{:ok, boolean}` if successful, `{:error, reason}` otherwise.
  """
  @spec get_table_totals_row(Spreadsheet.t(), String.t(), String.t()) ::
          {:ok, boolean()} | {:error, String.t()}
  def get_table_totals_row(%Spreadsheet{reference: resource}, sheet_name, table_name) do
    UmyaNative.get_table_totals_row(resource, sheet_name, table_name)
  end
end
