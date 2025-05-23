defmodule UmyaSpreadsheet.SheetIntegrationTest do
  use ExUnit.Case, async: true

  @output_path "test/result_files/sheet_integration_test.xlsx"

  setup do
    # Create a new spreadsheet for testing
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Clean up any previous test result files
    File.rm(@output_path)

    %{spreadsheet: spreadsheet}
  end

  test "clone and remove sheets", %{spreadsheet: spreadsheet} do
    # Add content to the default sheet
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Original Sheet")

    # Clone the sheet
    :ok = UmyaSpreadsheet.clone_sheet(spreadsheet, "Sheet1", "Sheet1 Clone")

    # Add a sheet that will be deleted
    :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "Temporary Sheet")

    # Get sheet names to verify our operations
    sheet_names = UmyaSpreadsheet.get_sheet_names(spreadsheet)
    assert Enum.member?(sheet_names, "Sheet1")
    assert Enum.member?(sheet_names, "Sheet1 Clone")
    assert Enum.member?(sheet_names, "Temporary Sheet")

    # Remove the temporary sheet
    :ok = UmyaSpreadsheet.remove_sheet(spreadsheet, "Temporary Sheet")

    # Verify the sheet was removed
    updated_sheet_names = UmyaSpreadsheet.get_sheet_names(spreadsheet)
    assert Enum.member?(updated_sheet_names, "Sheet1")
    assert Enum.member?(updated_sheet_names, "Sheet1 Clone")
    refute Enum.member?(updated_sheet_names, "Temporary Sheet")

    # Save the spreadsheet
    :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)

    # Read it back to verify it can be opened
    {:ok, read_spreadsheet} = UmyaSpreadsheet.read(@output_path)

    # Verify the sheet structure was preserved
    read_sheet_names = UmyaSpreadsheet.get_sheet_names(read_spreadsheet)
    assert Enum.member?(read_sheet_names, "Sheet1")
    assert Enum.member?(read_sheet_names, "Sheet1 Clone")
    refute Enum.member?(read_sheet_names, "Temporary Sheet")
  end

  test "insert and remove rows and columns", %{spreadsheet: spreadsheet} do
    # Add content to the default sheet
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Original Content")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Header B")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C1", "Header C")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "Row 2")

    # Insert new rows
    :ok = UmyaSpreadsheet.insert_new_row(spreadsheet, "Sheet1", 2, 3)

    # Add content to identify the inserted rows
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "Inserted Row 1")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A4", "Inserted Row 2")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A5", "Inserted Row 3")

    # Insert new columns
    :ok = UmyaSpreadsheet.insert_new_column(spreadsheet, "Sheet1", "B", 2)

    # Add content to identify the inserted columns
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Inserted Column 1")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C1", "Inserted Column 2")

    # Remove a row
    :ok = UmyaSpreadsheet.remove_row(spreadsheet, "Sheet1", 4, 1)

    # Remove a column by index
    :ok = UmyaSpreadsheet.remove_column_by_index(spreadsheet, "Sheet1", 3, 1)

    # Save the spreadsheet
    :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)

    # Read it back to verify it can be opened
    {:ok, _read_spreadsheet} = UmyaSpreadsheet.read(@output_path)
  end
end
