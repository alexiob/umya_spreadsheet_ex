defmodule RenameSheetTest do
  use ExUnit.Case, async: true

  test "rename_sheet functionality" do
    # Create a new spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Get initial sheet names
    initial_names = UmyaSpreadsheet.get_sheet_names(spreadsheet)
    assert "Sheet1" in initial_names

    # Test 1: Rename the default sheet
    assert :ok = UmyaSpreadsheet.rename_sheet(spreadsheet, "Sheet1", "Renamed Sheet")

    # Verify the rename worked
    updated_names = UmyaSpreadsheet.get_sheet_names(spreadsheet)
    assert "Renamed Sheet" in updated_names
    refute "Sheet1" in updated_names

    # Test 2: Try to rename to an existing name (should fail)
    :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "Another Sheet")

    assert {:error, _reason} =
             UmyaSpreadsheet.rename_sheet(spreadsheet, "Renamed Sheet", "Another Sheet")

    # Test 3: Try to rename a non-existent sheet (should fail)
    assert {:error, _reason} =
             UmyaSpreadsheet.rename_sheet(spreadsheet, "Non-existent", "New Name")

    # Final verification
    final_names = UmyaSpreadsheet.get_sheet_names(spreadsheet)
    assert "Renamed Sheet" in final_names
    assert "Another Sheet" in final_names
  end
end
