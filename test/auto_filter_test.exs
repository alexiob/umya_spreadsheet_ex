defmodule UmyaSpreadsheetTest.AutoFilterTest do
  use ExUnit.Case, async: true

  alias UmyaSpreadsheet

  setup do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()
    # Add some data to test auto filters
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Name")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Department")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C1", "Salary")

    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "John")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", "IT")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C2", "75000")

    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "Sarah")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B3", "HR")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C3", "65000")

    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A4", "Mike")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B4", "Sales")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C4", "85000")

    %{spreadsheet: spreadsheet}
  end

  test "set_auto_filter adds an auto filter to the specified range", %{spreadsheet: spreadsheet} do
    range = "A1:C4"
    assert :ok = UmyaSpreadsheet.set_auto_filter(spreadsheet, "Sheet1", range)

    # Check that the auto filter was set
    assert {:ok, true} = UmyaSpreadsheet.has_auto_filter(spreadsheet, "Sheet1")

    # Check that the filter range is correct
    assert {:ok, ^range} = UmyaSpreadsheet.get_auto_filter_range(spreadsheet, "Sheet1")
  end

  test "remove_auto_filter removes an auto filter from the worksheet", %{spreadsheet: spreadsheet} do
    # First set an auto filter
    :ok = UmyaSpreadsheet.set_auto_filter(spreadsheet, "Sheet1", "A1:C4")
    assert {:ok, true} = UmyaSpreadsheet.has_auto_filter(spreadsheet, "Sheet1")

    # Now remove it
    :ok = UmyaSpreadsheet.remove_auto_filter(spreadsheet, "Sheet1")

    # Check that it was removed
    assert {:ok, false} = UmyaSpreadsheet.has_auto_filter(spreadsheet, "Sheet1")
    assert {:ok, nil} = UmyaSpreadsheet.get_auto_filter_range(spreadsheet, "Sheet1")
  end

  test "get_auto_filter_range returns nil when no auto filter exists", %{spreadsheet: spreadsheet} do
    assert {:ok, nil} = UmyaSpreadsheet.get_auto_filter_range(spreadsheet, "Sheet1")
  end

  test "has_auto_filter returns false when no auto filter exists", %{spreadsheet: spreadsheet} do
    assert {:ok, false} = UmyaSpreadsheet.has_auto_filter(spreadsheet, "Sheet1")
  end

  test "set_auto_filter with invalid sheet name returns error", %{spreadsheet: spreadsheet} do
    assert {:error, _} = UmyaSpreadsheet.set_auto_filter(spreadsheet, "NonExistentSheet", "A1:C4")
  end

  test "auto filter persists when saving and loading the workbook", %{spreadsheet: spreadsheet} do
    # Create a temp file path
    temp_path = "test/result_files/auto_filter_test.xlsx"

    # Set an auto filter and save the file
    :ok = UmyaSpreadsheet.set_auto_filter(spreadsheet, "Sheet1", "A1:C4")
    :ok = UmyaSpreadsheet.write(spreadsheet, temp_path)

    # Load the file back
    {:ok, loaded_spreadsheet} = UmyaSpreadsheet.read(temp_path)

    # Check that the auto filter is still there
    assert {:ok, true} = UmyaSpreadsheet.has_auto_filter(loaded_spreadsheet, "Sheet1")
    assert {:ok, "A1:C4"} = UmyaSpreadsheet.get_auto_filter_range(loaded_spreadsheet, "Sheet1")

    # Clean up
    File.rm(temp_path)
  end
end
