defmodule UmyaSpreadsheet.FileFormatOptions do
  @moduledoc """
  Advanced file format options for Excel files.

  This module provides additional control over file format options that are available in the
  underlying Rust library but were not previously exposed in the Elixir wrapper.

  Options include:
  - Compression level control for XLSX files
  - Enhanced encryption options
  - Converting a spreadsheet directly to a binary without writing to disk
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaNative

  @doc """
  Writes a spreadsheet to disk with a specified compression level.

  Compression levels range from 0 (no compression) to 9 (maximum compression).
  Higher compression levels result in smaller files but take longer to create.

  ## Parameters

  * `spreadsheet` - The spreadsheet struct
  * `path` - Path where the Excel file will be saved
  * `compression_level` - Integer from 0 to 9 (0 = no compression, 9 = maximum compression)

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> UmyaSpreadsheet.FileFormatOptions.write_with_compression(spreadsheet, "high_compression.xlsx", 9)
      :ok

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> UmyaSpreadsheet.FileFormatOptions.write_with_compression(spreadsheet, "no_compression.xlsx", 0)
      :ok

  """
  @spec write_with_compression(Spreadsheet.t(), String.t(), non_neg_integer()) ::
          :ok | {:error, atom()}
  def write_with_compression(%Spreadsheet{reference: ref}, path, compression_level)
      when compression_level in 0..9 do
    case UmyaNative.write_with_compression(ref, path, compression_level) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      # Handle nested error tuples from Rust NIFs
      {:error, {:error, message}} -> {:error, message}
      {:error, :error} -> {:error, "Failed to write file with compression"}
      {:error, message} -> {:error, message}
      result -> result
    end
  end

  @doc """
  Writes a spreadsheet to disk with enhanced encryption options.

  This function provides more control over the encryption process than the standard
  `write_with_password` function, allowing you to specify encryption algorithm,
  salt values, and spin counts for enhanced security.

  ## Parameters

  * `spreadsheet` - The spreadsheet struct
  * `path` - Path where the encrypted Excel file will be saved
  * `password` - Password to protect the file with
  * `algorithm` - Encryption algorithm (e.g., "AES128", "AES256", "default")
  * `salt_value` - Optional salt value for password derivation (if nil, a random salt is used)
  * `spin_count` - Optional spin count for key derivation (if nil, Excel default is used)

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> UmyaSpreadsheet.FileFormatOptions.write_with_encryption_options(
      ...>   spreadsheet,
      ...>   "highly_secure.xlsx",
      ...>   "secret123",
      ...>   "AES256"
      ...> )
      :ok

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> UmyaSpreadsheet.FileFormatOptions.write_with_encryption_options(
      ...>   spreadsheet,
      ...>   "custom_encryption.xlsx",
      ...>   "very_secret",
      ...>   "AES256",
      ...>   "customSalt",
      ...>   100000
      ...> )
      :ok

  """
  @spec write_with_encryption_options(
          Spreadsheet.t(),
          String.t(),
          String.t(),
          String.t(),
          String.t() | nil,
          non_neg_integer() | nil
        ) :: :ok | {:error, atom()}
  def write_with_encryption_options(
        %Spreadsheet{reference: ref},
        path,
        password,
        algorithm,
        salt_value \\ nil,
        spin_count \\ nil
      ) do
    case UmyaNative.write_with_encryption_options(
           ref,
           path,
           password,
           algorithm,
           salt_value,
           spin_count
         ) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      # Handle nested error tuples from Rust NIFs
      {:error, {:error, message}} -> {:error, message}
      {:error, :error} -> {:error, "Failed to write file with encryption"}
      result -> result
    end
  end

  @doc """
  Converts a spreadsheet to binary XLSX format without writing to disk.

  This function is useful when you need to:
  - Send an Excel file in an HTTP response
  - Store Excel files in a database
  - Process Excel files in memory without touching the filesystem

  ## Parameters

  * `spreadsheet` - The spreadsheet struct

  ## Returns

  * Binary data representing the Excel file

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Hello")
      iex> binary = UmyaSpreadsheet.FileFormatOptions.to_binary_xlsx(spreadsheet)
      iex> is_binary(binary)
      true
      iex> byte_size(binary) > 0
      true

  """
  @spec to_binary_xlsx(Spreadsheet.t()) :: binary() | {:error, atom()}
  def to_binary_xlsx(%Spreadsheet{reference: ref}) do
    case UmyaNative.to_binary_xlsx(ref) do
      {:ok, binary} -> binary
      {:error, reason} -> {:error, reason}
      result -> result
    end
  end

  @doc """
  Gets the default compression level used for XLSX files.

  This function returns the current default compression level that will be used when
  writing XLSX files with the standard write function.

  ## Parameters

  * `spreadsheet` - The spreadsheet struct

  ## Returns

  * Integer from 0 to 9 representing the compression level (0 = no compression, 9 = maximum compression)

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> UmyaSpreadsheet.FileFormatOptions.get_compression_level(spreadsheet)
      6

  """
  @spec get_compression_level(Spreadsheet.t()) :: integer() | {:error, atom()}
  def get_compression_level(%Spreadsheet{reference: ref}) do
    case UmyaNative.get_compression_level(ref) do
      {:ok, level} -> level
      {:error, reason} -> {:error, reason}
      result -> result
    end
  end

  @doc """
  Checks if a spreadsheet has encryption enabled.

  This function checks if the spreadsheet has any workbook protection or encryption
  settings enabled.

  ## Parameters

  * `spreadsheet` - The spreadsheet struct

  ## Returns

  * Boolean indicating if encryption is enabled

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> UmyaSpreadsheet.FileFormatOptions.is_encrypted(spreadsheet)
      false

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.read_xlsx_with_password("encrypted.xlsx", "secret123")
      iex> UmyaSpreadsheet.FileFormatOptions.is_encrypted(spreadsheet)
      true

  """
  @spec is_encrypted(Spreadsheet.t()) :: boolean() | {:error, atom()}
  def is_encrypted(%Spreadsheet{reference: ref}) do
    case UmyaNative.is_encrypted(ref) do
      {:ok, encrypted} -> encrypted
      {:error, reason} -> {:error, reason}
      result -> result
    end
  end

  @doc """
  Gets the encryption algorithm used for a password-protected spreadsheet.

  This function returns the encryption algorithm used to protect the spreadsheet,
  or nil if the spreadsheet is not encrypted.

  ## Parameters

  * `spreadsheet` - The spreadsheet struct

  ## Returns

  * A string representing the encryption algorithm (e.g., "AES256") if the spreadsheet is encrypted
  * `nil` if the spreadsheet is not encrypted
  * `{:error, reason}` if the operation failed

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> UmyaSpreadsheet.FileFormatOptions.get_encryption_algorithm(spreadsheet)
      nil

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.read("encrypted.xlsx", "password")
      iex> UmyaSpreadsheet.FileFormatOptions.get_encryption_algorithm(spreadsheet)
      "AES256"
  """
  @spec get_encryption_algorithm(Spreadsheet.t()) :: String.t() | nil | {:error, atom()}
  def get_encryption_algorithm(%Spreadsheet{reference: ref}) do
    case UmyaNative.get_encryption_algorithm(ref) do
      {:ok, algorithm} -> algorithm
      {:error, reason} -> {:error, reason}
      result -> result
    end
  end
end
