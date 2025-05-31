# File Format Options

This guide covers the advanced file format options available in UmyaSpreadsheet, providing you with more control over XLSX file generation, compression, encryption, and delivery methods.

## Compression Control

Excel files (.xlsx) are essentially ZIP archives containing XML files. UmyaSpreadsheet allows you to control the compression level of these files, balancing between file size and generation speed.

### Compression Levels

Compression levels range from 0 (no compression) to 9 (maximum compression):

- **Level 0**: No compression, fastest creation time, largest file size
- **Level 1-3**: Light compression, good speed, moderate file size reduction
- **Level 4-6**: Balanced compression (Excel default is around level 6)
- **Level 7-9**: Maximum compression, slower creation time, smallest file size

### Getting the Default Compression Level

You can check what compression level is being used by default:

```elixir
alias UmyaSpreadsheet

{:ok, spreadsheet} = UmyaSpreadsheet.new()
level = UmyaSpreadsheet.FileFormatOptions.get_compression_level(spreadsheet)
IO.puts("Default compression level: #{level}")
# => Default compression level: 6
```

### Setting Custom Compression

```elixir
alias UmyaSpreadsheet

# Create a spreadsheet with a lot of data
{:ok, spreadsheet} = UmyaSpreadsheet.new()
# ... add data to the spreadsheet ...

# Save with no compression - fastest creation, largest file size
UmyaSpreadsheet.FileFormatOptions.write_with_compression(spreadsheet, "uncompressed.xlsx", 0)

# Save with default compression
UmyaSpreadsheet.write(spreadsheet, "default_compression.xlsx")

# Save with maximum compression - smallest file size, slower creation
UmyaSpreadsheet.FileFormatOptions.write_with_compression(spreadsheet, "max_compressed.xlsx", 9)
```

### When to Use Different Compression Levels

- **No compression (0)**: When generation speed is critical and file size doesn't matter
- **Light compression (1-3)**: For large files where you need a balance of speed and size
- **Default compression**: For most use cases
- **Maximum compression (7-9)**: When storage space or network bandwidth is limited

## Enhanced Encryption Options

UmyaSpreadsheet provides advanced encryption options for securing Excel files with passwords.

### Checking Encryption Status

You can check if a spreadsheet has encryption enabled:

```elixir
alias UmyaSpreadsheet

# Check if a spreadsheet is encrypted
{:ok, spreadsheet} = UmyaSpreadsheet.read_xlsx("document.xlsx")
if UmyaSpreadsheet.FileFormatOptions.is_encrypted(spreadsheet) do
  IO.puts("The spreadsheet has encryption enabled")
else
  IO.puts("The spreadsheet is not encrypted")
end

# Get the encryption algorithm used
algorithm = UmyaSpreadsheet.FileFormatOptions.get_encryption_algorithm(spreadsheet)
IO.puts("Encryption algorithm: #{algorithm || "None"}")
# => Encryption algorithm: AES256
```

### Basic Password Protection

For simple password protection, use the standard function:

```elixir
UmyaSpreadsheet.write_with_password(spreadsheet, "protected.xlsx", "myPassword")
```

### Advanced Encryption Options

For more control over the encryption process:

```elixir
# Use AES256 encryption with custom parameters
UmyaSpreadsheet.write_with_encryption_options(
  spreadsheet,
  "highly_secure.xlsx",
  "myPassword",
  "AES256",         # Algorithm
  "customSaltValue", # Optional salt value
  100000            # Optional spin count
)
```

### Available Encryption Algorithms

- **"default"**: Uses Excel's default encryption (typically AES128)
- **"AES128"**: 128-bit AES encryption
- **"AES256"**: 256-bit AES encryption (stronger than AES128)

### Security Considerations

- Using a custom salt value can make password cracking more difficult
- Higher spin counts increase the computational effort needed to decrypt files
- AES256 provides stronger encryption than AES128 but may not be necessary for all use cases

## Binary Excel Files

Sometimes you need to generate Excel files without writing them to disk, such as sending them directly in HTTP responses or storing them in a database.

### Converting a Spreadsheet to Binary

```elixir
# Create and populate a spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new()
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Hello World")

# Convert to binary
xlsx_binary = UmyaSpreadsheet.to_binary_xlsx(spreadsheet)
```

### Web Application Example (Phoenix)

```elixir
def download_report(conn, _params) do
  # Create spreadsheet
  {:ok, spreadsheet} = UmyaSpreadsheet.new()
  # ... add data to the spreadsheet ...

  # Convert to binary
  xlsx_data = UmyaSpreadsheet.to_binary_xlsx(spreadsheet)

  # Send as download
  conn
  |> put_resp_content_type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
  |> put_resp_header("content-disposition", ~s[attachment; filename="report.xlsx"])
  |> send_resp(200, xlsx_data)
end
```

## Choosing the Right Options

Here's a guide to help you choose the right file format options for different scenarios:

### For Large Spreadsheets

```elixir
# Use light writer with moderate compression
UmyaSpreadsheet.write_light(spreadsheet, "large_file.xlsx")

# Or with custom compression level
UmyaSpreadsheet.write_with_compression(spreadsheet, "large_file.xlsx", 4)
```

### For Secure Files

```elixir
# Basic password protection
UmyaSpreadsheet.write_with_password(spreadsheet, "secure.xlsx", "password123")

# More secure encryption
UmyaSpreadsheet.write_with_encryption_options(
  spreadsheet,
  "highly_secure.xlsx",
  "strongPassword",
  "AES256"
)
```

### For Web Applications

```elixir
# Generate binary for HTTP responses
xlsx_data = UmyaSpreadsheet.to_binary_xlsx(spreadsheet)

# Or compress it first (pseudo-code)
{:ok, spreadsheet} = UmyaSpreadsheet.new()
# ... add data ...
temp_path = "/tmp/temp_#{System.unique_integer()}.xlsx"
UmyaSpreadsheet.write_with_compression(spreadsheet, temp_path, 9)
xlsx_data = File.read!(temp_path)
File.rm!(temp_path)
```

## Performance Considerations

- Higher compression levels take more time to generate but produce smaller files
- The `write_light` function uses less memory but doesn't support all features
- Binary generation (`to_binary_xlsx`) keeps the entire file in memory
- For very large files, consider using CSV export or creating XLSX files with minimal formatting
