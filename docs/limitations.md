# UmyaSpreadsheet Limitations and Compatibility

This document outlines the current limitations, compatibility constraints, and known issues with UmyaSpreadsheet.

## Table of Contents

- [File Format Limitations](#file-format-limitations)
- [Feature Limitations](#feature-limitations)
- [Platform Compatibility](#platform-compatibility)
- [Version Compatibility](#version-compatibility)
- [Performance Limitations](#performance-limitations)
- [Known Issues](#known-issues)
- [Roadmap and Future Support](#roadmap-and-future-support)

## File Format Limitations

### Supported Formats

| Format | Read | Write | Notes |
|--------|------|-------|-------|
| .xlsx  | ✅   | ✅    | Full support |
| .xlsm  | ✅   | ✅    | Macro support limited |
| .xls   | ❌   | ❌    | Legacy format not supported |
| .csv   | ❌   | ✅    | Export only |
| .ods   | ❌   | ❌    | OpenDocument not supported |

### Excel Feature Support

| Feature | Support Level | Notes |
|---------|---------------|-------|
| **Basic Operations** |||
| Cell values (text, numbers) | ✅ Full | All basic data types |
| Formulas | ✅ Full | Standard Excel formulas |
| Multiple sheets | ✅ Full | Add, remove, rename sheets |
| **Formatting** |||
| Font styling | ✅ Full | Bold, italic, underline, colors |
| Cell backgrounds | ✅ Full | Colors and patterns |
| Borders | ✅ Full | All border styles |
| Number formats | ✅ Full | Currency, dates, percentages |
| Conditional formatting | ✅ Partial | Basic rules, color scales, data bars, icon sets |
| **Advanced Features** |||
| Charts | ✅ Partial | Line, bar, pie charts supported |
| Images | ✅ Full | PNG, JPEG insertion |
| Data validation | ✅ Full | Dropdowns, ranges, custom rules |
| Pivot tables | ✅ Basic | Creation and basic manipulation |
| Macros | ❌ Limited | Can preserve existing macros but not execute |
| **Security** |||
| Password protection | ✅ Full | Workbook and sheet protection |
| Digital signatures | ❌ None | Not supported |
| **Print Features** |||
| Page setup | ✅ Full | Margins, orientation, scaling |
| Headers/footers | ✅ Full | Custom headers and footers |
| Print areas | ✅ Full | Define print ranges |

### Unsupported Excel Features

- **Macros**: Cannot create or execute VBA macros (can preserve existing ones)
- **External Links**: Links to other files not fully supported
- **OLE Objects**: Embedded objects not supported
- **Smart Art**: Smart Art graphics not supported
- **3D Charts**: Only 2D chart types supported
- **Form Controls**: Excel form controls not supported
- **Custom Functions**: Excel add-in functions not supported
- **Collaboration Features**: Comments collaboration features limited
- **Power Query**: Power Query connections not supported

## Feature Limitations

### Formula Limitations

```elixir
# Supported formula types
UmyaSpreadsheet.set_cell_formula(spreadsheet, "Sheet1", "A1", "=SUM(B1:B10)")
UmyaSpreadsheet.set_array_formula(spreadsheet, "Sheet1", "A1:A10", "=B1:B10*2")

# Limitations:
# - Custom functions from add-ins not supported
# - Some advanced statistical functions may not work
# - External reference formulas limited
```

### Chart Limitations

```elixir
# Supported chart types
chart_types = [
  "line", "bar", "column", "pie", "area", "scatter", "doughnut"
]

# Not supported:
# - 3D charts
# - Bubble charts
# - Stock charts
# - Combo charts
# - Treemap charts
# - Waterfall charts
```

### Conditional Formatting Limitations

```elixir
# Supported rules
:ok = UmyaSpreadsheet.add_cell_value_rule(spreadsheet, sheet, range, operator, value, format)
:ok = UmyaSpreadsheet.add_color_scale(spreadsheet, sheet, range, colors)
:ok = UmyaSpreadsheet.add_data_bar(spreadsheet, sheet, range, color)
:ok = UmyaSpreadsheet.add_icon_set(spreadsheet, sheet, range, style, thresholds)

# Limitations:
# - Custom formatting rules not supported
# - Some advanced icon sets not available
# - Complex multi-condition rules limited
```

### Data Validation Limitations

```elixir
# Supported validation types
validation_types = [
  "list",        # Dropdown lists
  "whole",       # Whole numbers
  "decimal",     # Decimal numbers
  "date",        # Date ranges
  "time",        # Time ranges
  "text_length", # Text length constraints
  "custom"       # Custom formulas
]

# Limitations:
# - Error alert customization limited
# - Input message formatting limited
# - Circle invalid data feature not supported
```

## Platform Compatibility

### Operating Systems

| Platform | Support Level | Notes |
|----------|---------------|-------|
| **Linux** |||
| Ubuntu 20.04+ | ✅ Full | Precompiled binaries available |
| CentOS/RHEL 8+ | ✅ Full | Precompiled binaries available |
| Debian 10+ | ✅ Full | Precompiled binaries available |
| Alpine Linux | ⚠️ Limited | May require manual compilation |
| **macOS** |||
| macOS 11+ (Intel) | ✅ Full | Precompiled binaries available |
| macOS 11+ (Apple Silicon) | ✅ Full | Native ARM64 support |
| macOS 10.15 and earlier | ❌ None | Not supported |
| **Windows** |||
| Windows 10+ | ✅ Full | Precompiled binaries available |
| Windows Server 2019+ | ✅ Full | Precompiled binaries available |
| Windows 8.1 and earlier | ❌ None | Not supported |

### Architecture Support

| Architecture | Support Level | Notes |
|--------------|---------------|-------|
| x86_64 (Intel/AMD 64-bit) | ✅ Full | Primary support |
| aarch64 (ARM64) | ✅ Full | Apple Silicon, ARM servers |
| x86 (32-bit) | ❌ None | Not supported |
| ARM 32-bit | ❌ None | Not supported |

### Compilation Requirements

When building from source:

```bash
# Minimum Rust version
rustc --version  # Should be 1.70.0 or newer

# Required system dependencies (Linux)
sudo apt-get install build-essential pkg-config libssl-dev

# Required system dependencies (macOS)
# Xcode Command Line Tools
xcode-select --install

# Required system dependencies (Windows)
# Visual Studio Build Tools 2019 or newer
```

## Version Compatibility

### Elixir/OTP Compatibility Matrix

| UmyaSpreadsheet | Elixir | OTP | Notes |
|-----------------|--------|-----|-------|
| 0.6.x | 1.12+ | 24+ | Current version |
| 0.5.x | 1.11+ | 23+ | Legacy support |
| 0.4.x | 1.10+ | 22+ | End of life |

### Excel Compatibility

| Excel Version | Read Support | Write Support | Notes |
|---------------|--------------|---------------|-------|
| Excel 2019+ | ✅ Full | ✅ Full | Recommended |
| Excel 2016 | ✅ Full | ✅ Full | Fully supported |
| Excel 2013 | ✅ Full | ✅ Full | Fully supported |
| Excel 2010 | ✅ Full | ✅ Full | Fully supported |
| Excel 2007 | ✅ Full | ✅ Full | Minimum supported |
| Excel 2003 and earlier | ❌ None | ❌ None | Use .xls converter |

### LibreOffice/OpenOffice Compatibility

| Application | Read Support | Write Support | Notes |
|-------------|--------------|---------------|-------|
| LibreOffice Calc 7.x | ✅ Good | ✅ Good | Minor formatting differences |
| LibreOffice Calc 6.x | ✅ Good | ✅ Good | Minor formatting differences |
| OpenOffice Calc | ⚠️ Limited | ⚠️ Limited | Basic features only |

**Known LibreOffice Issues:**

- Some conditional formatting rules may not display correctly
- Chart formatting may differ slightly
- Custom number formats may not be preserved exactly

## Performance Limitations

### Memory Usage

```elixir
# Memory usage guidelines
file_size_mb = File.stat!("large_file.xlsx").size / (1024 * 1024)

memory_estimate = cond do
  file_size_mb < 10 -> "#{file_size_mb * 2-3} MB RAM"
  file_size_mb < 50 -> "#{file_size_mb * 3-4} MB RAM"
  file_size_mb < 100 -> "#{file_size_mb * 4-5} MB RAM"
  true -> "Consider using lazy loading or chunked processing"
end

IO.puts("Estimated memory usage: #{memory_estimate}")
```

### File Size Limits

| Operation | Recommended Limit | Hard Limit | Notes |
|-----------|-------------------|------------|-------|
| Read file | < 100 MB | ~ 500 MB | Use `lazy_read` for large files |
| Write file | < 100 MB | ~ 200 MB | Use `write_light` for simple data |
| In-memory operations | < 50 MB | ~ 100 MB | Depends on available RAM |
| Cell count | < 1M cells | ~ 10M cells | Performance degrades |

### Performance Benchmarks

Typical performance on modern hardware:

```elixir
# Approximate processing speeds
operations_per_second = %{
  cell_read: 50_000,          # cells/second
  cell_write: 25_000,         # cells/second
  file_read_10mb: 2,          # files/second
  file_write_10mb: 1,         # files/second
  formula_calculation: 10_000  # formulas/second
}
```

## Known Issues

### Issue: Global Defined Names Warning

**Symptom:**

```
Warning: Global defined name 'TaxRate' created without name due to API limitations in umya-spreadsheet 2.3.0
```

**Impact:** Cosmetic warning only, functionality works correctly

**Workaround:** None needed, can be safely ignored

**Status:** Will be fixed in future umya-spreadsheet release

### Issue: Password Protection Algorithm Limitations

**Symptom:**

```
{:error, "Failed to write file with password"}
```

**Impact:** Some advanced encryption algorithms not supported

**Workaround:**

```elixir
# Use basic password protection
UmyaSpreadsheet.write_with_password(spreadsheet, path, "password")

# For advanced encryption, use external tools
```

### Issue: Large File Memory Usage

**Symptom:** High memory usage with files > 50MB

**Impact:** Potential out-of-memory errors

**Workaround:**

```elixir
# Use lazy reading
{:ok, spreadsheet} = UmyaSpreadsheet.lazy_read("large_file.xlsx")

# Process in chunks
def process_large_file(file_path) do
  {:ok, spreadsheet} = UmyaSpreadsheet.lazy_read(file_path)

  # Process specific ranges instead of entire file
  for row <- 1..1000 do
    value = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "A#{row}")
    process_value(value)
  end
end
```

### Issue: Concurrent Access Limitations

**Symptom:** Crashes when multiple processes access same spreadsheet reference

**Impact:** Race conditions in multi-process scenarios

**Workaround:**

```elixir
# Use separate spreadsheet instances per process
Task.async(fn ->
  {:ok, spreadsheet} = UmyaSpreadsheet.new()
  # ... work with spreadsheet
end)

# Or coordinate access via GenServer
defmodule SpreadsheetServer do
  use GenServer
  # ... implementation
end
```

## Roadmap and Future Support

### Planned Features (Future Releases)

- **Enhanced Chart Support**: More chart types and customization options
- **Improved Macro Support**: Better VBA macro preservation and analysis
- **Advanced Pivot Tables**: More pivot table features and calculations
- **External Data Sources**: Support for external data connections
- **Collaboration Features**: Better support for shared workbooks
- **Performance Improvements**: Streaming operations for very large files

### Breaking Changes Policy

UmyaSpreadsheet follows semantic versioning:

- **Major versions** (1.0, 2.0): Breaking API changes
- **Minor versions** (0.6.0, 0.7.0): New features, backward compatible
- **Patch versions** (0.6.1, 0.6.2): Bug fixes only

### Migration Guides

When upgrading between major versions, migration guides will be provided to help update existing code.

### Long-term Support

- Current major version: Maintained with bug fixes and security updates
- Previous major version: Bug fixes only for 6 months after new major release
- Older versions: Community support only

## Getting Support

For issues related to limitations:

1. **Check this document** for known limitations
2. **Review GitHub Issues** for existing reports
3. **Check upstream issues** in the Rust umya-spreadsheet library
4. **Submit feature requests** with clear use cases

### Contribution Guidelines

Help improve UmyaSpreadsheet by:

1. **Reporting bugs** with minimal reproduction cases
2. **Suggesting features** with detailed requirements
3. **Contributing code** to address limitations
4. **Improving documentation** with real-world examples

For more information, see the [DEVELOPMENT.md](../DEVELOPMENT.md) guide.
