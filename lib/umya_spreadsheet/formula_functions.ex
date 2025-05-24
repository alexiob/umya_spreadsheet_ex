defmodule UmyaSpreadsheet.FormulaFunctions do
  @moduledoc """
  Functions for working with formulas in spreadsheets.
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaNative

  @doc """
  Sets a regular formula in a cell.

  ## Parameters

  * `spreadsheet` - The spreadsheet struct
  * `sheet_name` - Name of the worksheet
  * `cell_address` - Cell address in A1 notation (e.g., "A1", "B2")
  * `formula` - Formula string (without leading =)

  ## Examples

      iex> UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "A1", "SUM(B1:B10)")
      :ok

  """
  @spec set_formula(Spreadsheet.t(), String.t(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_formula(%Spreadsheet{reference: ref}, sheet_name, cell_address, formula) do
    UmyaNative.set_formula(ref, sheet_name, cell_address, formula)
  end

  @doc """
  Sets an array formula for a range of cells. Array formulas can return multiple values
  across a range of cells.

  ## Parameters

  * `spreadsheet` - The spreadsheet struct
  * `sheet_name` - Name of the worksheet
  * `range` - Cell range in A1 notation (e.g., "A1:B5")
  * `formula` - Formula string (without leading =)

  ## Examples

      iex> UmyaSpreadsheet.set_array_formula(spreadsheet, "Sheet1", "A1:A3", "ROW(1:3)")
      :ok

  """
  @spec set_array_formula(Spreadsheet.t(), String.t(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_array_formula(%Spreadsheet{reference: ref}, sheet_name, range, formula) do
    UmyaNative.set_array_formula(ref, sheet_name, range, formula)
  end

  @doc """
  Creates a named range in the spreadsheet.

  ## Parameters

  * `spreadsheet` - The spreadsheet struct
  * `name` - Name for the range
  * `sheet_name` - Name of the worksheet
  * `range` - Cell range in A1 notation (e.g., "A1:B5")

  ## Examples

      iex> UmyaSpreadsheet.create_named_range(spreadsheet, "MyRange", "Sheet1", "A1:B10")
      :ok

  """
  @spec create_named_range(Spreadsheet.t(), String.t(), String.t(), String.t()) :: :ok | {:error, atom()}
  def create_named_range(%Spreadsheet{reference: ref}, name, sheet_name, range) do
    UmyaNative.create_named_range(ref, name, sheet_name, range)
  end

  @doc """
  Creates a defined name in the spreadsheet with an associated formula.

  ## Parameters

  * `spreadsheet` - The spreadsheet struct
  * `name` - Name for the defined name
  * `formula` - Formula string (without leading =)
  * `sheet_name` - Optional name of the worksheet to scope the defined name to

  ## Examples

      iex> UmyaSpreadsheet.create_defined_name(spreadsheet, "TaxRate", "0.15")
      :ok

      iex> UmyaSpreadsheet.create_defined_name(spreadsheet, "Department", "Sales", "Sheet1")
      :ok

  """
  @spec create_defined_name(Spreadsheet.t(), String.t(), String.t(), String.t() | nil) :: :ok | {:error, atom()}
  def create_defined_name(%Spreadsheet{reference: ref}, name, formula, sheet_name \\ nil) do
    UmyaNative.create_defined_name(ref, name, formula, sheet_name)
  end

  @doc """
  Gets all defined names in the spreadsheet.

  ## Parameters

  * `spreadsheet` - The spreadsheet struct

  ## Returns

  * List of tuples containing name and formula/address

  ## Examples

      iex> UmyaSpreadsheet.get_defined_names(spreadsheet)
      [{"MyRange", "Sheet1!A1:B10"}, {"TaxRate", "0.15"}]

  """
  @spec get_defined_names(Spreadsheet.t()) :: [{String.t(), String.t()}] | {:error, atom()}
  def get_defined_names(%Spreadsheet{reference: ref}) do
    UmyaNative.get_defined_names(ref)
  end
end
