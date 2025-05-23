defmodule UmyaSpreadsheet.SheetOperationsTest do
  use ExUnit.Case, async: true

  @output_path "test/result_files/sheet_operations_output.xlsx"

  setup do
    # Create a new spreadsheet for testing
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Clean up any previous test result files
    File.rm(@output_path)

    %{spreadsheet: spreadsheet}
  end

  test "clone and remove sheets", %{spreadsheet: spreadsheet} do
    # First verify the default sheet is present
    sheet_names = UmyaSpreadsheet.get_sheet_names(spreadsheet)
    assert "Sheet1" in sheet_names

    # Clone the sheet with a new name
    assert :ok = UmyaSpreadsheet.clone_sheet(spreadsheet, "Sheet1", "ClonedSheet")

    # Verify the new sheet is in the sheet list
    sheet_names = UmyaSpreadsheet.get_sheet_names(spreadsheet)
    assert "ClonedSheet" in sheet_names

    # Remove the cloned sheet
    assert :ok = UmyaSpreadsheet.remove_sheet(spreadsheet, "ClonedSheet")

    # Verify the sheet was removed
    sheet_names = UmyaSpreadsheet.get_sheet_names(spreadsheet)
    refute "ClonedSheet" in sheet_names
  end

  test "insert and remove rows and columns", %{spreadsheet: spreadsheet} do
    # Prepare some data
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Row 1, Col A")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Row 1, Col B")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "Row 2, Col A")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", "Row 2, Col B")

    # Insert a new row at index 2
    assert :ok = UmyaSpreadsheet.insert_new_row(spreadsheet, "Sheet1", 2, 1)

    # Verify row was inserted
    assert {:ok, "Row 1, Col A"} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "A1")
    assert {:ok, ""} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "A2")
    assert {:ok, "Row 2, Col A"} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "A3")

    # Insert a new column at column B
    assert :ok = UmyaSpreadsheet.insert_new_column(spreadsheet, "Sheet1", "B", 1)

    # Verify column was inserted
    assert {:ok, "Row 1, Col A"} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "A1")
    assert {:ok, ""} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "B1")
    assert {:ok, "Row 1, Col B"} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "C1")

    # Remove the inserted row
    assert :ok = UmyaSpreadsheet.remove_row(spreadsheet, "Sheet1", 2, 1)

    # Verify row was removed
    assert {:ok, "Row 1, Col A"} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "A1")
    assert {:ok, "Row 2, Col A"} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "A2")

    # Remove the inserted column
    assert :ok = UmyaSpreadsheet.remove_column(spreadsheet, "Sheet1", "B", 1)

    # Verify column was removed
    assert {:ok, "Row 1, Col A"} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "A1")
    assert {:ok, "Row 1, Col B"} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "B1")

    # Save the spreadsheet for manual verification if needed
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)
  end

  test "set font name for a cell", %{spreadsheet: spreadsheet} do
    # Set cell value and font name
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Custom Font Cell")
    assert :ok = UmyaSpreadsheet.set_font_name(spreadsheet, "Sheet1", "A1", "Arial")

    # Save the spreadsheet for manual verification
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)
  end
end
