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

  test "get_pivot_table_names returns list of pivot table names", %{spreadsheet: spreadsheet} do
    # Initially no pivot tables
    assert {:ok, []} = PivotTable.get_pivot_table_names(spreadsheet, "PivotSheet")

    # Create a pivot table
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

    # Now should return the name
    assert {:ok, ["Sales Analysis"]} = PivotTable.get_pivot_table_names(spreadsheet, "PivotSheet")

    # Add another pivot table
    :ok =
      PivotTable.add_pivot_table(
        spreadsheet,
        "PivotSheet",
        "Product Analysis",
        "Sheet1",
        "A1:D5",
        "E3",
        [1],
        [],
        [{2, "sum", "Total Sales"}]
      )

    # Should return both names
    assert {:ok, names} = PivotTable.get_pivot_table_names(spreadsheet, "PivotSheet")
    assert length(names) == 2
    assert "Sales Analysis" in names
    assert "Product Analysis" in names
  end

  test "get_pivot_table_info returns detailed information", %{spreadsheet: spreadsheet} do
    # Create a pivot table
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

    # Get detailed info
    assert {:ok, {name, location, source_range, cache_id}} =
             PivotTable.get_pivot_table_info(spreadsheet, "PivotSheet", "Sales Analysis")

    assert name == "Sales Analysis"
    assert is_binary(location)
    assert is_binary(source_range)
    assert is_binary(cache_id) || is_integer(cache_id)

    # Test non-existent pivot table
    assert {:error, :not_found} =
             PivotTable.get_pivot_table_info(spreadsheet, "PivotSheet", "Non-Existent")
  end

  test "get_pivot_table_source_range returns source information", %{spreadsheet: spreadsheet} do
    # Create a pivot table
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

    # Get source range
    assert {:ok, {source_sheet, source_range}} =
             PivotTable.get_pivot_table_source_range(spreadsheet, "PivotSheet", "Sales Analysis")

    assert source_sheet == "Sheet1"
    assert source_range == "A1:D5"

    # Test non-existent pivot table
    assert {:error, :not_found} =
             PivotTable.get_pivot_table_source_range(spreadsheet, "PivotSheet", "Non-Existent")
  end

  test "get_pivot_table_target_cell returns target cell location", %{spreadsheet: spreadsheet} do
    # Create a pivot table
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

    # Get target cell
    assert {:ok, target_cell} =
             PivotTable.get_pivot_table_target_cell(spreadsheet, "PivotSheet", "Sales Analysis")

    # Should start with "A3" (may have offset appended)
    assert String.starts_with?(target_cell, "A3")

    # Test non-existent pivot table
    assert {:error, :not_found} =
             PivotTable.get_pivot_table_target_cell(spreadsheet, "PivotSheet", "Non-Existent")
  end

  test "get_pivot_table_fields returns field configuration", %{spreadsheet: spreadsheet} do
    # Create a pivot table with specific field configuration
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

    # Get field configuration
    assert {:ok, {row_fields, column_fields, data_fields}} =
             PivotTable.get_pivot_table_fields(spreadsheet, "PivotSheet", "Sales Analysis")

    assert is_list(row_fields)
    assert is_list(column_fields)
    assert is_list(data_fields)

    # Verify field configuration matches what we set
    assert 0 in row_fields
    assert 1 in column_fields
    assert length(data_fields) == 1

    # Test non-existent pivot table
    assert {:error, :not_found} =
             PivotTable.get_pivot_table_fields(spreadsheet, "PivotSheet", "Non-Existent")
  end

  test "get functions handle non-existent sheet", %{spreadsheet: spreadsheet} do
    # Test all getter functions with non-existent sheet
    assert {:error, :not_found} =
             PivotTable.get_pivot_table_names(spreadsheet, "NonExistentSheet")

    assert {:error, :not_found} =
             PivotTable.get_pivot_table_info(spreadsheet, "NonExistentSheet", "Any Table")

    assert {:error, :not_found} =
             PivotTable.get_pivot_table_source_range(spreadsheet, "NonExistentSheet", "Any Table")

    assert {:error, :not_found} =
             PivotTable.get_pivot_table_target_cell(spreadsheet, "NonExistentSheet", "Any Table")

    assert {:error, :not_found} =
             PivotTable.get_pivot_table_fields(spreadsheet, "NonExistentSheet", "Any Table")
  end

  test "multiple pivot tables getter functions", %{spreadsheet: spreadsheet} do
    # Create multiple pivot tables with different configurations
    :ok =
      PivotTable.add_pivot_table(
        spreadsheet,
        "PivotSheet",
        "Region Analysis",
        "Sheet1",
        "A1:D5",
        "A3",
        [0],
        [],
        [{2, "sum", "Total Sales"}]
      )

    :ok =
      PivotTable.add_pivot_table(
        spreadsheet,
        "PivotSheet",
        "Product Analysis",
        "Sheet1",
        "A1:D5",
        "E3",
        [1],
        [0],
        [{2, "average", "Avg Sales"}]
      )

    # Test get_pivot_table_names
    assert {:ok, names} = PivotTable.get_pivot_table_names(spreadsheet, "PivotSheet")
    assert length(names) == 2
    assert "Region Analysis" in names
    assert "Product Analysis" in names

    # Test get_pivot_table_info for both
    assert {:ok, {name1, _, _, _}} =
             PivotTable.get_pivot_table_info(spreadsheet, "PivotSheet", "Region Analysis")

    assert {:ok, {name2, _, _, _}} =
             PivotTable.get_pivot_table_info(spreadsheet, "PivotSheet", "Product Analysis")

    assert name1 == "Region Analysis"
    assert name2 == "Product Analysis"

    # Test source ranges (should be the same for both)
    assert {:ok, {source_sheet1, source_range1}} =
             PivotTable.get_pivot_table_source_range(spreadsheet, "PivotSheet", "Region Analysis")

    assert {:ok, {source_sheet2, source_range2}} =
             PivotTable.get_pivot_table_source_range(
               spreadsheet,
               "PivotSheet",
               "Product Analysis"
             )

    assert source_sheet1 == source_sheet2
    assert source_range1 == source_range2

    # Test field configurations (should be different)
    assert {:ok, {row_fields1, column_fields1, _data_fields1}} =
             PivotTable.get_pivot_table_fields(spreadsheet, "PivotSheet", "Region Analysis")

    assert {:ok, {row_fields2, column_fields2, _data_fields2}} =
             PivotTable.get_pivot_table_fields(spreadsheet, "PivotSheet", "Product Analysis")

    # First pivot table has region (0) as row field, no column fields
    assert 0 in row_fields1
    assert column_fields1 == []

    # Second pivot table has product (1) as row field, region (0) as column field
    assert 1 in row_fields2
    assert 0 in column_fields2
  end
end
