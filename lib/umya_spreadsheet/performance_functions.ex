defmodule UmyaSpreadsheet.PerformanceFunctions do
  @moduledoc """
  Functions for optimized file writing with reduced memory usage.
  """
  
  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaNative

  @doc """
  Writes a spreadsheet to a file with reduced memory consumption.
  
  This variant is suitable for very large spreadsheets that would otherwise
  exceed memory limits when written with the standard `write/2` function.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `path` - Path where the file should be written

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.write_light(spreadsheet, "large_file.xlsx")
  """
  def write_light(%Spreadsheet{reference: ref}, path) do
    case UmyaNative.write_file_light(ref, path) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Writes a password-protected spreadsheet to a file with reduced memory consumption.
  
  This variant is suitable for very large spreadsheets that would otherwise
  exceed memory limits when written with the standard `write_with_password/3` function.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `path` - Path where the file should be written
  - `password` - Password to protect the file with

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.write_with_password_light(spreadsheet, "protected_large_file.xlsx", "secret123")
  """
  def write_with_password_light(%Spreadsheet{reference: ref}, path, password) do
    case UmyaNative.write_file_with_password_light(ref, path, password) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end
end
