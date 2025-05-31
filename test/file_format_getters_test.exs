defmodule UmyaSpreadsheetTest.FileFormatGettersTest do
  use ExUnit.Case, async: true

  alias UmyaSpreadsheet

  setup do
    # Create a standard spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Test data")
    
    %{spreadsheet: spreadsheet}
  end

  test "getter functions exist and return expected types", %{spreadsheet: spreadsheet} do
    # Test that get_compression_level returns an integer
    level = UmyaSpreadsheet.FileFormatOptions.get_compression_level(spreadsheet)
    assert is_integer(level)
    assert level == 6  # Default compression level is 6
    
    # Test that is_encrypted returns a boolean
    encrypted = UmyaSpreadsheet.FileFormatOptions.is_encrypted(spreadsheet)
    assert is_boolean(encrypted)
    
    # Test that get_encryption_algorithm returns nil for unencrypted spreadsheet
    algorithm = UmyaSpreadsheet.FileFormatOptions.get_encryption_algorithm(spreadsheet)
    assert algorithm == nil
  end
end
