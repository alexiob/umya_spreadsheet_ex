# Workbook Protection

This guide covers the workbook protection functionality in UmyaSpreadsheet, allowing you to check and manage workbook protection settings.

## Checking Workbook Protection

You can check if a workbook has protection enabled:

```elixir
alias UmyaSpreadsheet

# Open an existing spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.read("protected_document.xlsx")

# Check if the workbook has protection enabled
case UmyaSpreadsheet.is_workbook_protected(spreadsheet) do
  {:ok, true} -> IO.puts("This workbook has protection enabled")
  {:ok, false} -> IO.puts("This workbook is not protected")
  {:error, reason} -> IO.puts("Error checking protection: #{reason}")
end
```

## Getting Protection Details

You can retrieve detailed information about the workbook protection settings:

```elixir
# Get protection details
case UmyaSpreadsheet.get_workbook_protection_details(spreadsheet) do
  {:ok, protection_details} ->
    # Check specific protection settings
    if protection_details["lock_structure"] == "true" do
      IO.puts("The workbook structure is locked (cannot add/remove/rename sheets)")
    end

    if protection_details["lock_windows"] == "true" do
      IO.puts("The workbook windows are locked")
    end

    if protection_details["lock_revision"] == "true" do
      IO.puts("Revision tracking is locked")
    end
  {:error, reason} ->
    IO.puts("Error getting protection details: #{reason}")
end
```

## Working with Password Protected Files

When you need to read or write password-protected files:

```elixir
# Read a password-protected file
{:ok, spreadsheet} = UmyaSpreadsheet.read_with_password("protected_workbook.xlsx", "secretpassword123")

# Write a file with password protection
UmyaSpreadsheet.write_with_password(spreadsheet, "new_protected_workbook.xlsx", "newpassword456")
```

## Protection Settings vs. Password Protection

It's important to understand the difference between:

1. **Workbook protection settings** - Controls what users can do within the workbook (modify structure, windows, etc.)
2. **Password protection** - Controls access to open or modify the file itself

Use `is_workbook_protected/1` and `get_workbook_protection_details/1` to check the first type, which relates to structural and interface protections rather than file access controls.

## Best Practices for Workbook Protection

1. **Document Intent**: When using protection, document why specific protections are in place
2. **Balance Security and Usability**: Apply only the protection settings needed for your use case
3. **Consider Password Strength**: If using password protection, use strong passwords
4. **Multi-level Protection**: Consider both workbook and worksheet level protection for sensitive data
