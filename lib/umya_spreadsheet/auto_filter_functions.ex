defmodule UmyaSpreadsheet.AutoFilterFunctions do
  @moduledoc """
  Functions for working with auto filters in spreadsheets.

  Auto filters allow users to filter data in Excel by adding dropdown menus to column headers.
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaNative

  @doc """
  Sets an auto filter for a range of cells in a worksheet.

  This adds filter dropdown buttons to the top row of the specified range.

  ## Parameters

  * `spreadsheet` - The spreadsheet struct
  * `sheet_name` - Name of the worksheet
  * `range` - Cell range in A1 notation (e.g., "A1:E10")

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> UmyaSpreadsheet.set_auto_filter(spreadsheet, "Sheet1", "A1:E10")
      :ok

  """
  @spec set_auto_filter(Spreadsheet.t(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_auto_filter(%Spreadsheet{reference: ref}, sheet_name, range) do
    UmyaNative.set_auto_filter(ref, sheet_name, range)
  end

  @doc """
  Removes an auto filter from a worksheet.

  ## Parameters

  * `spreadsheet` - The spreadsheet struct
  * `sheet_name` - Name of the worksheet

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> UmyaSpreadsheet.set_auto_filter(spreadsheet, "Sheet1", "A1:E10")
      iex> UmyaSpreadsheet.remove_auto_filter(spreadsheet, "Sheet1")
      :ok

  """
  @spec remove_auto_filter(Spreadsheet.t(), String.t()) :: :ok | {:error, atom()}
  def remove_auto_filter(%Spreadsheet{reference: ref}, sheet_name) do
    UmyaNative.remove_auto_filter(ref, sheet_name)
  end

  @doc """
  Checks if a worksheet has an auto filter.

  ## Parameters

  * `spreadsheet` - The spreadsheet struct
  * `sheet_name` - Name of the worksheet

  ## Returns

  * `true` if the worksheet has an auto filter, `false` otherwise

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> UmyaSpreadsheet.set_auto_filter(spreadsheet, "Sheet1", "A1:E10")
      iex> UmyaSpreadsheet.has_auto_filter(spreadsheet, "Sheet1")
      true

  """
  @spec has_auto_filter(Spreadsheet.t(), String.t()) :: boolean() | {:error, atom()}
  def has_auto_filter(%Spreadsheet{reference: ref}, sheet_name) do
    UmyaNative.has_auto_filter(ref, sheet_name)
  end

  @doc """
  Gets the range of an auto filter in a worksheet.

  ## Parameters

  * `spreadsheet` - The spreadsheet struct
  * `sheet_name` - Name of the worksheet

  ## Returns

  * The range of the auto filter (e.g., "A1:E10") or `nil` if no auto filter exists

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> UmyaSpreadsheet.set_auto_filter(spreadsheet, "Sheet1", "A1:E10")
      iex> UmyaSpreadsheet.get_auto_filter_range(spreadsheet, "Sheet1")
      "A1:E10"

  """
  @spec get_auto_filter_range(Spreadsheet.t(), String.t()) :: String.t() | nil | {:error, atom()}
  def get_auto_filter_range(%Spreadsheet{reference: ref}, sheet_name) do
    UmyaNative.get_auto_filter_range(ref, sheet_name)
  end
end
