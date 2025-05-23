defmodule UmyaSpreadsheet.CSVFunctions do
  @moduledoc """
  Functions for exporting spreadsheets to CSV format.
  """
  
  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaNative

  @doc """
  Exports a sheet to CSV format with default options.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet to export
  - `path` - The output path for the CSV file

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.CSVFunctions.write_csv(spreadsheet, "Sheet1", "output.csv")
  """
  def write_csv(%Spreadsheet{reference: ref}, sheet_name, path) do
    case UmyaNative.write_csv(ref, sheet_name, path) do
      {:ok, :ok} -> :ok
      result -> result
    end
  end

  @doc """
  Exports a sheet to CSV format with custom options.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet to export
  - `path` - The output path for the CSV file
  - `options` - A map of CSV writer options with the following fields:
      - `:encoding` - The character encoding to use (default: "UTF8")
      - `:delimiter` - The field delimiter (default: ",")
      - `:do_trim` - Whether to trim whitespace (default: false)
      - `:wrap_with_char` - Character to wrap fields with (default: "")

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      
      options = %{
        encoding: "UTF8",
        delimiter: ";", 
        do_trim: true,
        wrap_with_char: "\""
      }
      
      :ok = UmyaSpreadsheet.CSVFunctions.write_csv_with_options(spreadsheet, "Sheet1", "output.csv", options)
  """
  def write_csv_with_options(%Spreadsheet{reference: ref}, sheet_name, path, options) do
    # Set default values
    encoding = Map.get(options, :encoding, "UTF8")
    delimiter = Map.get(options, :delimiter, ",")
    do_trim = Map.get(options, :do_trim, false)
    wrap_with_char = Map.get(options, :wrap_with_char, "")

    case UmyaNative.write_csv_with_options(
      ref,
      sheet_name,
      path,
      encoding,
      delimiter,
      do_trim,
      wrap_with_char
    ) do
      {:ok, :ok} -> :ok
      result -> result
    end
  end

  @doc """
  Exports a sheet to CSV format with custom options.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet to export
  - `path` - The output path for the CSV file
  - `options` - A map of CSV writer options with the following fields:
      - `:encoding` - The character encoding to use (default: "UTF8")
      - `:delimiter` - The field delimiter (default: ",")
      - `:do_trim` - Whether to trim whitespace (default: false)
      - `:wrap_with_char` - Character to wrap fields with (default: "")

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      
      options = %{
        encoding: "UTF8",
        delimiter: ";", 
        do_trim: true,
        wrap_with_char: "\""
      }
      
      :ok = UmyaSpreadsheet.CSVFunctions.write_csv(spreadsheet, "Sheet1", "output.csv", options)
  """
  def write_csv(%Spreadsheet{reference: ref}, sheet_name, path, options) do
    # For compatibility with write_csv_with_options
    write_csv_with_options(%Spreadsheet{reference: ref}, sheet_name, path, options)
  end
end
