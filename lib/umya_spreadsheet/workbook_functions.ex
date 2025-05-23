defmodule UmyaSpreadsheet.WorkbookFunctions do
  @moduledoc """
  Functions for manipulating workbook properties and security.
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaNative

  @doc """
  Sets workbook protection with a password.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `password` - The password to protect the workbook with

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.WorkbookFunctions.set_workbook_protection(spreadsheet, "password123")
  """
  def set_workbook_protection(%Spreadsheet{reference: ref}, password) do
    case UmyaNative.set_workbook_protection(ref, password) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets password protection for an existing Excel file.

  This function adds password protection to an existing file without loading it into memory.

  ## Parameters

  - `input_path` - Path to the input Excel file
  - `output_path` - Path where the protected file should be saved
  - `password` - The password to protect the file with

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      :ok = UmyaSpreadsheet.WorkbookFunctions.set_password("input.xlsx", "protected.xlsx", "password123")
  """
  def set_password(input_path, output_path, password) do
    case UmyaNative.set_password(input_path, output_path, password) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end
end
