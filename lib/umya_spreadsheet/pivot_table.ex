defmodule UmyaSpreadsheet.PivotTable do
  @moduledoc """
  Functions for creating and manipulating pivot tables.

  Pivot tables allow you to summarize and analyze large datasets quickly.
  This module provides functions to:

  * Create pivot tables from data ranges
  * Configure row, column, and data fields
  * Apply formatting to pivot tables
  * Refresh pivot table data
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaNative

  @doc """
  Adds a new pivot table to a spreadsheet.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet where the pivot table will be placed
    * `name` - Name of the pivot table
    * `source_sheet` - Name of the sheet containing the source data
    * `source_range` - Range of cells containing the source data in A1 notation (e.g., "A1:D10")
    * `target_cell` - Top-left cell for the pivot table placement
    * `row_fields` - List of field indices (0-based) to use as row fields
    * `column_fields` - List of field indices (0-based) to use as column fields
    * `data_fields` - List of data field configs in the format [{field_index, "Function", "Custom Name"}]
      where function is one of "sum", "count", "average", "max", "min", "product", "count_nums", "stddev", "stddevp", "var", "varp"

  ## Examples

  ```elixir
  # Create a simple pivot table from data in sheet "Data"
  PivotTable.add_pivot_table(
    spreadsheet,
    "PivotSheet",
    "Sales Analysis",
    "Data",
    "A1:D100",
    "A3",
    [0], # Use first column (Region) as row field
    [1], # Use second column (Product) as column field
    [{2, "sum", "Total Sales"}] # Sum the third column (Sales) as data field
  )
  ```
  """
  @spec add_pivot_table(
          Spreadsheet.t(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          [integer()],
          [integer()],
          [{integer(), String.t(), String.t()}]
        ) :: :ok | {:error, atom()}
  def add_pivot_table(
        %Spreadsheet{reference: ref},
        sheet_name,
        name,
        source_sheet,
        source_range,
        target_cell,
        row_fields,
        column_fields,
        data_fields
      ) do
    case UmyaNative.add_pivot_table(
           UmyaSpreadsheet.unwrap_ref(ref),
           sheet_name,
           name,
           source_sheet,
           source_range,
           target_cell,
           row_fields,
           column_fields,
           data_fields
         ) do
      :ok -> :ok
      {:ok, :ok} -> :ok
      {:error, reason} -> {:error, reason}
    end
  end

  @doc """
  Checks if a sheet contains any pivot tables.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet to check

  ## Examples

  ```elixir
  # Check if a sheet has pivot tables
  if PivotTable.has_pivot_tables?(spreadsheet, "Sheet1") do
    # Handle sheet with pivot tables
  end
  ```
  """
  @spec has_pivot_tables?(Spreadsheet.t(), String.t()) :: boolean()
  def has_pivot_tables?(%Spreadsheet{reference: ref}, sheet_name) do
    UmyaNative.has_pivot_tables(UmyaSpreadsheet.unwrap_ref(ref), sheet_name)
  end

  @doc """
  Gets the number of pivot tables in a sheet.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet to check

  ## Examples

  ```elixir
  # Get number of pivot tables
  count = PivotTable.count_pivot_tables(spreadsheet, "Sheet1")
  ```
  """
  @spec count_pivot_tables(Spreadsheet.t(), String.t()) :: integer()
  def count_pivot_tables(%Spreadsheet{reference: ref}, sheet_name) do
    case UmyaNative.count_pivot_tables(UmyaSpreadsheet.unwrap_ref(ref), sheet_name) do
      {:ok, count} when is_integer(count) -> count
      count when is_integer(count) -> count
      _ -> 0
    end
  end

  @doc """
  Refreshes all pivot tables in a spreadsheet.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct

  ## Examples

  ```elixir
  # Refresh all pivot tables
  PivotTable.refresh_all_pivot_tables(spreadsheet)
  ```
  """
  @spec refresh_all_pivot_tables(Spreadsheet.t()) :: :ok | {:error, atom()}
  def refresh_all_pivot_tables(%Spreadsheet{reference: ref}) do
    case UmyaNative.refresh_all_pivot_tables(UmyaSpreadsheet.unwrap_ref(ref)) do
      :ok -> :ok
      {:ok, :ok} -> :ok
      {:error, reason} -> {:error, reason}
    end
  end

  @doc """
  Removes a pivot table from a sheet by name.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet containing the pivot table
    * `pivot_table_name` - Name of the pivot table to remove

  ## Examples

  ```elixir
  # Remove a specific pivot table
  PivotTable.remove_pivot_table(spreadsheet, "Sheet1", "Sales Analysis")
  ```
  """
  @spec remove_pivot_table(Spreadsheet.t(), String.t(), String.t()) :: :ok | {:error, atom()}
  def remove_pivot_table(%Spreadsheet{reference: ref}, sheet_name, pivot_table_name) do
    case UmyaNative.remove_pivot_table(
           UmyaSpreadsheet.unwrap_ref(ref),
           sheet_name,
           pivot_table_name
         ) do
      :ok -> :ok
      {:ok, :ok} -> :ok
      {:error, reason} -> {:error, reason}
    end
  end

  @doc """
  Gets a list of all pivot table names in a sheet.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet to check

  ## Returns

    * `{:ok, [String.t()]}` - List of pivot table names
    * `{:error, atom()}` - Error if sheet doesn't exist

  ## Examples

      # Get all pivot table names in a sheet
      case PivotTable.get_pivot_table_names(spreadsheet, "Sheet1") do
        {:ok, names} -> IO.inspect(names)
        {:error, reason} -> IO.puts("Error: " <> inspect(reason))
      end

  """
  @spec get_pivot_table_names(Spreadsheet.t(), String.t()) ::
          {:ok, [String.t()]} | {:error, atom()}
  def get_pivot_table_names(%Spreadsheet{reference: ref}, sheet_name) do
    UmyaNative.get_pivot_table_names(UmyaSpreadsheet.unwrap_ref(ref), sheet_name)
  end

  @doc """
  Gets detailed information about a specific pivot table.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet containing the pivot table
    * `pivot_table_name` - Name of the pivot table

  ## Returns

    * `{:ok, {name, location, source_range, cache_id}}` - Tuple containing pivot table details
    * `{:error, atom()}` - Error if sheet or pivot table doesn't exist

  ## Examples

      # Get detailed info about a pivot table
      case PivotTable.get_pivot_table_info(spreadsheet, "Sheet1", "Sales Analysis") do
        {:ok, {name, location, source_range, cache_id}} ->
          IO.puts("Name: " <> name <> ", Location: " <> location)
        {:error, reason} ->
          IO.puts("Error: " <> inspect(reason))
      end

  """
  @spec get_pivot_table_info(Spreadsheet.t(), String.t(), String.t()) ::
          {:ok, {String.t(), String.t(), String.t(), String.t()}} | {:error, atom()}
  def get_pivot_table_info(%Spreadsheet{reference: ref}, sheet_name, pivot_table_name) do
    UmyaNative.get_pivot_table_info(UmyaSpreadsheet.unwrap_ref(ref), sheet_name, pivot_table_name)
  end

  @doc """
  Gets the source data range for a pivot table.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet containing the pivot table
    * `pivot_table_name` - Name of the pivot table

  ## Returns

    * `{:ok, {source_sheet, source_range}}` - Tuple containing source sheet name and range
    * `{:error, atom()}` - Error if sheet or pivot table doesn't exist

  ## Examples

      # Get source range for a pivot table
      case PivotTable.get_pivot_table_source_range(spreadsheet, "Sheet1", "Sales Analysis") do
        {:ok, {source_sheet, source_range}} ->
          IO.puts("Source: " <> source_sheet <> "!" <> source_range)
        {:error, reason} ->
          IO.puts("Error: " <> inspect(reason))
      end

  """
  @spec get_pivot_table_source_range(Spreadsheet.t(), String.t(), String.t()) ::
          {:ok, {String.t(), String.t()}} | {:error, atom()}
  def get_pivot_table_source_range(%Spreadsheet{reference: ref}, sheet_name, pivot_table_name) do
    UmyaNative.get_pivot_table_source_range(
      UmyaSpreadsheet.unwrap_ref(ref),
      sheet_name,
      pivot_table_name
    )
  end

  @doc """
  Gets the target cell (top-left placement) for a pivot table.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet containing the pivot table
    * `pivot_table_name` - Name of the pivot table

  ## Returns

    * `{:ok, String.t()}` - Target cell reference (e.g., "A3")
    * `{:error, atom()}` - Error if sheet or pivot table doesn't exist

  ## Examples

      # Get target cell for a pivot table
      case PivotTable.get_pivot_table_target_cell(spreadsheet, "Sheet1", "Sales Analysis") do
        {:ok, target_cell} ->
          IO.puts("Pivot table starts at: " <> target_cell)
        {:error, reason} ->
          IO.puts("Error: " <> inspect(reason))
      end

  """
  @spec get_pivot_table_target_cell(Spreadsheet.t(), String.t(), String.t()) ::
          {:ok, String.t()} | {:error, atom()}
  def get_pivot_table_target_cell(%Spreadsheet{reference: ref}, sheet_name, pivot_table_name) do
    UmyaNative.get_pivot_table_target_cell(
      UmyaSpreadsheet.unwrap_ref(ref),
      sheet_name,
      pivot_table_name
    )
  end

  @doc """
  Gets the field configuration for a pivot table.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet containing the pivot table
    * `pivot_table_name` - Name of the pivot table

  ## Returns

    * `{:ok, {row_fields, column_fields, data_fields}}` - Tuple containing field configurations
      * `row_fields` - List of field indices used as row fields
      * `column_fields` - List of field indices used as column fields
      * `data_fields` - List of data field configs in format [{field_index, function, custom_name}]
    * `{:error, atom()}` - Error if sheet or pivot table doesn't exist

  ## Examples

      # Get field configuration for a pivot table
      case PivotTable.get_pivot_table_fields(spreadsheet, "Sheet1", "Sales Analysis") do
        {:ok, {row_fields, column_fields, data_fields}} ->
          IO.puts("Row fields: " <> inspect(row_fields))
          IO.puts("Column fields: " <> inspect(column_fields))
          IO.puts("Data fields: " <> inspect(data_fields))
        {:error, reason} ->
          IO.puts("Error: " <> inspect(reason))
      end

  """
  @spec get_pivot_table_fields(Spreadsheet.t(), String.t(), String.t()) ::
          {:ok, {[integer()], [integer()], [{integer(), String.t(), String.t()}]}}
          | {:error, atom()}
  def get_pivot_table_fields(%Spreadsheet{reference: ref}, sheet_name, pivot_table_name) do
    UmyaNative.get_pivot_table_fields(
      UmyaSpreadsheet.unwrap_ref(ref),
      sheet_name,
      pivot_table_name
    )
  end
end
