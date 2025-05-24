#!/usr/bin/env elixir

# Load the library
Mix.install([{:umya_spreadsheet_ex, path: "."}])

alias UmyaSpreadsheet

IO.puts("Testing create_defined_name function...")

# Create a new spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Test global defined name
IO.puts("Testing global defined name...")
result = UmyaSpreadsheet.create_defined_name(spreadsheet, "GLOBAL_CONST", "42", nil)
IO.inspect(result, label: "Global defined name result")

# Add a new sheet
UmyaSpreadsheet.add_sheet(spreadsheet, "TestSheet")

# Test sheet-scoped defined name
IO.puts("Testing sheet-scoped defined name...")
result2 = UmyaSpreadsheet.create_defined_name(spreadsheet, "LOCAL_CALC", "10+5", "TestSheet")
IO.inspect(result2, label: "Sheet-scoped defined name result")

# Test error cases
IO.puts("Testing error cases...")

# Empty name
result3 = UmyaSpreadsheet.create_defined_name(spreadsheet, "", "A1", nil)
IO.inspect(result3, label: "Empty name result")

# Empty formula
result4 = UmyaSpreadsheet.create_defined_name(spreadsheet, "TEST", "", nil)
IO.inspect(result4, label: "Empty formula result")

# Invalid sheet
result5 = UmyaSpreadsheet.create_defined_name(spreadsheet, "TEST", "A1", "NonExistent")
IO.inspect(result5, label: "Invalid sheet result")

# Try to save the file to verify it works
output_path = "test_defined_names.xlsx"
save_result = UmyaSpreadsheet.write(spreadsheet, output_path)
IO.inspect(save_result, label: "Save result")

if save_result == :ok do
  IO.puts("SUCCESS: Spreadsheet saved successfully!")

  # Try to read it back
  {:ok, loaded} = UmyaSpreadsheet.read(output_path)
  IO.puts("SUCCESS: Spreadsheet loaded successfully!")

  # Clean up
  File.rm(output_path)
else
  IO.puts("ERROR: Failed to save spreadsheet")
end

IO.puts("Testing completed!")
