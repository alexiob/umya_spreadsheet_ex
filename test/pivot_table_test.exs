defmodule UmyaSpreadsheetTest.PivotTableTest do
  use ExUnit.Case, async: true

  alias UmyaSpreadsheet
  alias UmyaSpreadsheet.PivotTable

  setup do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Create sample data in Sheet1
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Region")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Product")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C1", "Sales")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D1", "Date")

    # Row 2
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "North")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", "Apples")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C2", "10000")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D2", "2025-01-15")

    # Row 3
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "North")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B3", "Oranges")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C3", "8000")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D3", "2025-01-20")

    # Row 4
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A4", "South")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B4", "Apples")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C4", "12000")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D4", "2025-02-10")

    # Row 5
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A5", "South")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B5", "Oranges")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C5", "9000")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D5", "2025-02-15")

    # Add a second sheet for the pivot table
    UmyaSpreadsheet.add_sheet(spreadsheet, "PivotSheet")

    %{spreadsheet: spreadsheet}
  end

  test "add_pivot_table creates a new pivot table", %{spreadsheet: spreadsheet} do
    # Create a simple pivot table
    assert :ok =
             PivotTable.add_pivot_table(
               spreadsheet,
               "PivotSheet",
               "Sales Analysis",
               "Sheet1",
               "A1:D5",
               "A3",
               # Use Region (first column) as row field
               [0],
               # Use Product (second column) as column field
               [1],
               # Sum the Sales (third column)
               [{2, "sum", "Total Sales"}]
             )

    # Check that we have a pivot table now
    assert PivotTable.has_pivot_tables?(spreadsheet, "PivotSheet")
    assert 1 == PivotTable.count_pivot_tables(spreadsheet, "PivotSheet")

    # We could add file save/read test here, but that would be complex
    # For now, just verify that the API functions work without errors
  end

  test "remove_pivot_table removes a pivot table", %{spreadsheet: spreadsheet} do
    # First create a pivot table
    :ok =
      PivotTable.add_pivot_table(
        spreadsheet,
        "PivotSheet",
        "Sales Analysis",
        "Sheet1",
        "A1:D5",
        "A3",
        [0],
        [1],
        [{2, "sum", "Total Sales"}]
      )

    # Check that it was created
    assert PivotTable.has_pivot_tables?(spreadsheet, "PivotSheet")

    # Now remove it
    assert :ok = PivotTable.remove_pivot_table(spreadsheet, "PivotSheet", "Sales Analysis")

    # Check that it was removed
    refute PivotTable.has_pivot_tables?(spreadsheet, "PivotSheet")
    assert 0 == PivotTable.count_pivot_tables(spreadsheet, "PivotSheet")
  end

  test "multiple pivot tables on same sheet", %{spreadsheet: spreadsheet} do
    # Create first pivot table
    :ok =
      PivotTable.add_pivot_table(
        spreadsheet,
        "PivotSheet",
        "Region Analysis",
        "Sheet1",
        "A1:D5",
        "A3",
        # Region as row
        [0],
        # No column fields
        [],
        [{2, "sum", "Total Sales"}]
      )

    # Create second pivot table
    :ok =
      PivotTable.add_pivot_table(
        spreadsheet,
        "PivotSheet",
        "Product Analysis",
        "Sheet1",
        "A1:D5",
        "E3",
        # Product as row
        [1],
        # No column fields
        [],
        [{2, "sum", "Total Sales"}]
      )

    # Check that both were created
    assert PivotTable.has_pivot_tables?(spreadsheet, "PivotSheet")
    assert 2 == PivotTable.count_pivot_tables(spreadsheet, "PivotSheet")

    # Remove one
    assert :ok = PivotTable.remove_pivot_table(spreadsheet, "PivotSheet", "Region Analysis")

    # Check that we still have one left
    assert PivotTable.has_pivot_tables?(spreadsheet, "PivotSheet")
    assert 1 == PivotTable.count_pivot_tables(spreadsheet, "PivotSheet")
  end

  test "refresh_all_pivot_tables refreshes pivot tables", %{spreadsheet: spreadsheet} do
    # First create a pivot table
    :ok =
      PivotTable.add_pivot_table(
        spreadsheet,
        "PivotSheet",
        "Sales Analysis",
        "Sheet1",
        "A1:D5",
        "A3",
        [0],
        [1],
        [{2, "sum", "Total Sales"}]
      )

    # Now refresh all pivot tables
    # Note: This is mostly a smoke test since actual refreshing is
    # handled internally by the Rust library
    assert :ok = PivotTable.refresh_all_pivot_tables(spreadsheet)
  end
end
