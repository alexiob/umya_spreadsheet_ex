# UmyaSpreadsheet Troubleshooting Guide

This guide covers common issues, their solutions, and best practices for using UmyaSpreadsheet effectively.

## Table of Contents

- [Installation Issues](#installation-issues)
- [Runtime Errors](#runtime-errors)
- [Performance Issues](#performance-issues)
- [File Format Issues](#file-format-issues)
- [Memory Issues](#memory-issues)
- [Concurrency Issues](#concurrency-issues)
- [Platform-Specific Issues](#platform-specific-issues)
- [Best Practices](#best-practices)
- [Debugging Tips](#debugging-tips)

## Installation Issues

### Problem: NIF Compilation Fails

**Symptoms:**
```
** (Mix) could not compile dependency :umya_spreadsheet_ex
```

**Solutions:**

1. **Ensure Rust is installed:**
   ```bash
   curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh
   source ~/.cargo/env
   ```

2. **Force rebuild:**
   ```bash
   export UMYA_SPREADSHEET_BUILD=true
   mix deps.clean umya_spreadsheet_ex --build
   mix deps.get
   mix deps.compile
   ```

3. **Check Rust version compatibility:**
   ```bash
   rustc --version  # Should be 1.70.0 or newer
   ```

### Problem: Precompiled NIF Not Found

**Symptoms:**
```
** (RuntimeError) Unable to load NIF
```

**Solutions:**

1. **Use force build mode:**
   ```elixir
   # In config/config.exs
   config :umya_spreadsheet_ex, force_build: true
   ```

2. **Check platform compatibility:**
   ```bash
   mix deps.get
   mix deps.compile --force
   ```

### Problem: Apple Silicon (M1/M2) Issues

**Symptoms:**
```
** (ArgumentError) errors were found at the given arguments
```

**Solutions:**

1. **Install Rust for Apple Silicon:**
   ```bash
   arch -arm64 brew install rust
   ```

2. **Force ARM64 compilation:**
   ```bash
   export CARGO_TARGET_AARCH64_APPLE_DARWIN_LINKER=clang
   export UMYA_SPREADSHEET_BUILD=true
   mix deps.compile --force
   ```

## Runtime Errors

### Problem: Spreadsheet Reference Errors

**Symptoms:**
```
** (ArgumentError) argument error
```

**Common Causes & Solutions:**

1. **Using invalid spreadsheet reference:**
   ```elixir
   # Wrong - using a closed or invalid reference
   {:ok, spreadsheet} = UmyaSpreadsheet.new()
   # ... reference becomes invalid somehow ...
   UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "test")
   
   # Solution - always check if operations succeed
   case UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "test") do
     :ok -> :ok
     {:error, reason} -> 
       Logger.error("Failed to set cell value: #{reason}")
       # Handle error appropriately
   end
   ```

2. **Concurrent access to the same reference:**
   ```elixir
   # Wrong - sharing reference between processes
   Task.async(fn -> UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "1") end)
   Task.async(fn -> UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "2") end)
   
   # Solution - use separate spreadsheet instances
   Task.async(fn -> 
     {:ok, s1} = UmyaSpreadsheet.new()
     UmyaSpreadsheet.set_cell_value(s1, "Sheet1", "A1", "1")
   end)
   ```

### Problem: Sheet Not Found Errors

**Symptoms:**
```
{:error, "Sheet not found"}
```

**Solutions:**

1. **Check sheet names:**
   ```elixir
   # List all sheets first
   sheets = UmyaSpreadsheet.get_sheet_names(spreadsheet)
   IO.inspect(sheets)  # ["Sheet1", "Sheet2", ...]
   
   # Use exact sheet name
   UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "value")
   ```

2. **Handle missing sheets gracefully:**
   ```elixir
   def safe_set_cell_value(spreadsheet, sheet_name, cell, value) do
     case UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, cell, value) do
       :ok -> :ok
       {:error, "Sheet not found"} ->
         UmyaSpreadsheet.add_sheet(spreadsheet, sheet_name)
         UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, cell, value)
       {:error, reason} -> {:error, reason}
     end
   end
   ```

### Problem: Invalid Cell References

**Symptoms:**
```
{:error, "Invalid cell reference"}
```

**Solutions:**

1. **Use valid Excel cell references:**
   ```elixir
   # Valid formats
   UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "value")
   UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "Z100", "value")
   UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "AA1", "value")
   
   # Invalid formats (will fail)
   # UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "1A", "value")
   # UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A", "value")
   ```

2. **Validate cell references:**
   ```elixir
   def validate_cell_ref(cell_ref) do
     case Regex.match?(~r/^[A-Z]+[1-9]\d*$/, cell_ref) do
       true -> :ok
       false -> {:error, "Invalid cell reference format"}
     end
   end
   ```

## Performance Issues

### Problem: Slow File Operations

**Symptoms:**
- Long write times for large files
- High memory usage during operations

**Solutions:**

1. **Use light writers for simple operations:**
   ```elixir
   # For simple spreadsheets without complex formatting
   UmyaSpreadsheet.write_light(spreadsheet, "output.xlsx")
   ```

2. **Use lazy reading for large files:**
   ```elixir
   # For reading large files
   {:ok, spreadsheet} = UmyaSpreadsheet.lazy_read("large_file.xlsx")
   ```

3. **Batch operations:**
   ```elixir
   # Instead of individual cell operations
   for row <- 1..1000 do
     UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A#{row}", "value#{row}")
   end
   
   # Use bulk operations when available
   data = for row <- 1..1000, do: {"A#{row}", "value#{row}"}
   # Set multiple values in batch (if such function exists)
   ```

### Problem: Memory Usage Growing

**Symptoms:**
- Increasing memory usage over time
- Out of memory errors with large files

**Solutions:**

1. **Use streaming operations:**
   ```elixir
   # Process data in chunks
   def process_large_file(file_path) do
     {:ok, spreadsheet} = UmyaSpreadsheet.lazy_read(file_path)
     
     # Process data in smaller chunks
     Enum.chunk_every(1..10000, 100)
     |> Enum.each(fn chunk ->
       process_chunk(spreadsheet, chunk)
       # Allow garbage collection
       :erlang.garbage_collect()
     end)
   end
   ```

2. **Dispose of references when done:**
   ```elixir
   def process_and_cleanup() do
     {:ok, spreadsheet} = UmyaSpreadsheet.new()
     
     try do
       # Do work with spreadsheet
       UmyaSpreadsheet.write(spreadsheet, "output.xlsx")
     after
       # Spreadsheet reference will be garbage collected
       # when it goes out of scope
       :ok
     end
   end
   ```

## File Format Issues

### Problem: Corrupted Output Files

**Symptoms:**
- Excel reports file corruption
- Unable to open generated files

**Solutions:**

1. **Ensure proper file extensions:**
   ```elixir
   # Use .xlsx for Excel files
   UmyaSpreadsheet.write(spreadsheet, "output.xlsx")
   
   # Not .xls (older format not supported)
   ```

2. **Check file permissions:**
   ```elixir
   def safe_write(spreadsheet, path) do
     case UmyaSpreadsheet.write(spreadsheet, path) do
       :ok -> :ok
       {:error, reason} -> 
         Logger.error("Failed to write file #{path}: #{reason}")
         {:error, reason}
     end
   end
   ```

3. **Validate data before writing:**
   ```elixir
   def validate_and_write(spreadsheet, path) do
     # Check if spreadsheet has valid data
     case UmyaSpreadsheet.get_sheet_names(spreadsheet) do
       [] -> {:error, "No sheets in spreadsheet"}
       _sheets -> UmyaSpreadsheet.write(spreadsheet, path)
     end
   end
   ```

### Problem: Encoding Issues

**Symptoms:**
- Special characters not displaying correctly
- UTF-8 encoding problems

**Solutions:**

1. **Ensure UTF-8 encoding:**
   ```elixir
   # When setting cell values with special characters
   UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Café ñoño 中文")
   ```

2. **Handle CSV encoding properly:**
   ```elixir
   # When exporting to CSV
   UmyaSpreadsheet.to_csv(spreadsheet, "Sheet1", "output.csv")
   ```

## Memory Issues

### Problem: Out of Memory Errors

**Symptoms:**
```
** (SystemLimitError) a system limit has been reached
```

**Solutions:**

1. **Process data in smaller chunks:**
   ```elixir
   def process_large_dataset(data) do
     data
     |> Enum.chunk_every(1000)  # Process 1000 rows at a time
     |> Enum.reduce({:ok, _} = UmyaSpreadsheet.new(), fn chunk, {:ok, spreadsheet} ->
       process_chunk(spreadsheet, chunk)
       {:ok, spreadsheet}
     end)
   end
   ```

2. **Use memory-efficient patterns:**
   ```elixir
   # Instead of keeping all data in memory
   def stream_to_excel(data_stream, output_path) do
     {:ok, spreadsheet} = UmyaSpreadsheet.new()
     
     data_stream
     |> Stream.with_index(1)
     |> Stream.each(fn {row_data, row_num} ->
       row_data
       |> Enum.with_index(1)
       |> Enum.each(fn {value, col_num} ->
         cell = "#{<<64 + col_num>>}#{row_num}"
         UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", cell, value)
       end)
     end)
     |> Stream.run()
     
     UmyaSpreadsheet.write(spreadsheet, output_path)
   end
   ```

## Concurrency Issues

### Problem: Race Conditions

**Symptoms:**
- Inconsistent results with concurrent operations
- Occasional crashes with multiple processes

**Solutions:**

1. **Use separate spreadsheet instances:**
   ```elixir
   # Good - each task has its own spreadsheet
   tasks = for i <- 1..10 do
     Task.async(fn ->
       {:ok, spreadsheet} = UmyaSpreadsheet.new()
       UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Task #{i}")
       UmyaSpreadsheet.write(spreadsheet, "output_#{i}.xlsx")
     end)
   end
   
   Task.await_many(tasks)
   ```

2. **Coordinate access with GenServer:**
   ```elixir
   defmodule SpreadsheetManager do
     use GenServer
     
     def start_link(opts) do
       GenServer.start_link(__MODULE__, opts, name: __MODULE__)
     end
     
     def update_cell(sheet_name, cell, value) do
       GenServer.call(__MODULE__, {:update_cell, sheet_name, cell, value})
     end
     
     def init(_opts) do
       {:ok, spreadsheet} = UmyaSpreadsheet.new()
       {:ok, %{spreadsheet: spreadsheet}}
     end
     
     def handle_call({:update_cell, sheet_name, cell, value}, _from, state) do
       result = UmyaSpreadsheet.set_cell_value(state.spreadsheet, sheet_name, cell, value)
       {:reply, result, state}
     end
   end
   ```

## Platform-Specific Issues

### macOS Issues

**Problem: Code signing issues**
```bash
# Solution: Allow unsigned binaries (development only)
sudo spctl --master-disable
```

**Problem: Rosetta issues on Apple Silicon**
```bash
# Solution: Use native ARM64 Rust
arch -arm64 brew install rust
export UMYA_SPREADSHEET_BUILD=true
mix deps.compile --force
```

### Linux Issues

**Problem: Missing system dependencies**
```bash
# Solution: Install required packages
sudo apt-get update
sudo apt-get install build-essential pkg-config libssl-dev
```

### Windows Issues

**Problem: Visual Studio Build Tools not found**
```cmd
REM Solution: Install Visual Studio Build Tools
winget install Microsoft.VisualStudio.2022.BuildTools
```

## Best Practices

### 1. Error Handling

Always handle errors appropriately:

```elixir
def safe_spreadsheet_operation(spreadsheet, operation) do
  case operation.(spreadsheet) do
    :ok -> :ok
    {:ok, result} -> {:ok, result}
    {:error, reason} -> 
      Logger.error("Spreadsheet operation failed: #{reason}")
      {:error, reason}
  end
end
```

### 2. Resource Management

Create spreadsheets close to where they're used:

```elixir
def generate_report(data) do
  {:ok, spreadsheet} = UmyaSpreadsheet.new()
  
  try do
    populate_spreadsheet(spreadsheet, data)
    UmyaSpreadsheet.write(spreadsheet, "report.xlsx")
  rescue
    error -> 
      Logger.error("Failed to generate report: #{inspect(error)}")
      {:error, "Report generation failed"}
  end
end
```

### 3. Performance Optimization

Use appropriate functions for your use case:

```elixir
# For simple operations
UmyaSpreadsheet.write_light(spreadsheet, path)

# For large files
{:ok, spreadsheet} = UmyaSpreadsheet.lazy_read(path)

# For memory optimization
UmyaSpreadsheet.to_binary_xlsx(spreadsheet)
```

### 4. Data Validation

Validate inputs before processing:

```elixir
def validate_cell_data(cell_ref, value) do
  with :ok <- validate_cell_reference(cell_ref),
       :ok <- validate_cell_value(value) do
    :ok
  else
    {:error, reason} -> {:error, reason}
  end
end

def validate_cell_reference(cell_ref) when is_binary(cell_ref) do
  if Regex.match?(~r/^[A-Z]+[1-9]\d*$/, cell_ref) do
    :ok
  else
    {:error, "Invalid cell reference format"}
  end
end

def validate_cell_value(value) when is_binary(value) or is_number(value), do: :ok
def validate_cell_value(_), do: {:error, "Invalid cell value type"}
```

## Debugging Tips

### 1. Enable Debug Logging

```elixir
# In config/config.exs
config :logger, level: :debug

# In your code
require Logger
Logger.debug("Spreadsheet operation: #{inspect(operation_details)}")
```

### 2. Inspect Spreadsheet State

```elixir
def debug_spreadsheet(spreadsheet) do
  sheets = UmyaSpreadsheet.get_sheet_names(spreadsheet)
  Logger.debug("Available sheets: #{inspect(sheets)}")
  
  for sheet <- sheets do
    Logger.debug("Sheet #{sheet} active: #{UmyaSpreadsheet.is_sheet_visible(spreadsheet, sheet)}")
  end
end
```

### 3. Test with Minimal Examples

```elixir
def minimal_test() do
  {:ok, spreadsheet} = UmyaSpreadsheet.new()
  :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "test")
  value = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "A1")
  IO.puts("Cell A1 value: #{value}")
  UmyaSpreadsheet.write(spreadsheet, "/tmp/test.xlsx")
end
```

### 4. Check System Resources

```elixir
# Monitor memory usage
:observer.start()

# Check process memory
:erlang.memory()

# Monitor file handles
System.cmd("lsof", ["-p", "#{System.otp_release()}"])
```

## Getting Help

If you encounter issues not covered in this guide:

1. **Check the GitHub Issues**: [umya_spreadsheet_ex issues](https://github.com/alexiob/umya_spreadsheet_ex/issues)
2. **Review the Documentation**: [HexDocs](https://hexdocs.pm/umya_spreadsheet_ex/)
3. **Check Rust Library Issues**: [umya-spreadsheet issues](https://github.com/MathNya/umya-spreadsheet/issues)
4. **Create a Minimal Reproduction**: Provide the smallest possible code that reproduces the issue

When reporting issues, include:
- Elixir and OTP versions
- Operating system and architecture
- UmyaSpreadsheet version
- Minimal code to reproduce the issue
- Full error messages and stack traces
