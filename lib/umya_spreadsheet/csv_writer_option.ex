defmodule UmyaSpreadsheet.CsvWriterOption do
  @moduledoc """
  Represents configuration options for CSV export.

  This module provides a way to configure options for CSV export, including:
  - Encoding format (UTF-8, ShiftJIS, etc.)
  - Trimming options for cell values
  - Character wrapping for cell values
  - Delimiter character for separating values
  """

  defstruct csv_encode_value: :utf8,
            do_trim: false,
            wrap_with_char: "",
            delimiter: ","

  @type t :: %__MODULE__{
          csv_encode_value: atom(),
          do_trim: boolean(),
          wrap_with_char: String.t(),
          delimiter: String.t()
        }

  @doc """
  Creates a new CSV writer option struct with default values.

  ## Examples

      iex> UmyaSpreadsheet.CsvWriterOption.new()
      %UmyaSpreadsheet.CsvWriterOption{
        csv_encode_value: :utf8,
        do_trim: false,
        wrap_with_char: "",
        delimiter: ","
      }
  """
  def new do
    %__MODULE__{}
  end

  @doc """
  Sets the CSV encoding value.

  Available encodings:
    - `:utf8` - UTF-8 encoding (default)
    - `:shift_jis` - Shift JIS encoding
    - `:koi8u` - KOI8-U encoding
    - `:koi8r` - KOI8-R encoding
    - `:iso88598i` - ISO-8859-8-I encoding
    - `:gbk` - GBK encoding

  ## Examples

      iex> opts = UmyaSpreadsheet.CsvWriterOption.new()
      iex> UmyaSpreadsheet.CsvWriterOption.set_csv_encode_value(opts, :shift_jis)
      %UmyaSpreadsheet.CsvWriterOption{csv_encode_value: :shift_jis, do_trim: false, wrap_with_char: ""}
  """
  def set_csv_encode_value(%__MODULE__{} = options, encode_value)
      when encode_value in [:utf8, :shift_jis, :koi8u, :koi8r, :iso88598i, :gbk] do
    %__MODULE__{options | csv_encode_value: encode_value}
  end

  @doc """
  Sets whether to trim whitespace from cell values.

  ## Examples

      iex> opts = UmyaSpreadsheet.CsvWriterOption.new()
      iex> UmyaSpreadsheet.CsvWriterOption.set_do_trim(opts, true)
      %UmyaSpreadsheet.CsvWriterOption{csv_encode_value: :utf8, do_trim: true, wrap_with_char: ""}
  """
  def set_do_trim(%__MODULE__{} = options, do_trim) when is_boolean(do_trim) do
    %__MODULE__{options | do_trim: do_trim}
  end

  @doc """
  Sets the character to wrap cell values with.

  ## Examples

      iex> opts = UmyaSpreadsheet.CsvWriterOption.new()
      iex> UmyaSpreadsheet.CsvWriterOption.set_wrap_with_char(opts, "\"")
      %UmyaSpreadsheet.CsvWriterOption{csv_encode_value: :utf8, do_trim: false, wrap_with_char: "\\\"", delimiter: ","}
  """
  def set_wrap_with_char(%__MODULE__{} = options, wrap_with_char)
      when is_binary(wrap_with_char) do
    %__MODULE__{options | wrap_with_char: wrap_with_char}
  end

  @doc """
  Sets the delimiter character used to separate fields in the CSV file.
  Default is comma (",").

  ## Examples

      iex> opts = UmyaSpreadsheet.CsvWriterOption.new()
      iex> UmyaSpreadsheet.CsvWriterOption.set_delimiter(opts, ";")
      %UmyaSpreadsheet.CsvWriterOption{csv_encode_value: :utf8, do_trim: false, wrap_with_char: "", delimiter: ";"}
  """
  def set_delimiter(%__MODULE__{} = options, delimiter) when is_binary(delimiter) do
    %__MODULE__{options | delimiter: delimiter}
  end

  @doc """
  Converts an Elixir encoding atom to the Rust encoding value.

  This function is for internal use to convert our friendly atom names to the exact
  encoding names used in the Rust library.
  """
  def encode_value_to_string(:utf8), do: "UTF8"
  def encode_value_to_string(:shift_jis), do: "ShiftJis"
  def encode_value_to_string(:koi8u), do: "Koi8u"
  def encode_value_to_string(:koi8r), do: "Koi8r"
  def encode_value_to_string(:iso88598i), do: "Iso88598i"
  def encode_value_to_string(:gbk), do: "Gbk"
end
