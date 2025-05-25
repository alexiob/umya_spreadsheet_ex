defmodule UmyaSpreadsheetTest.TableTest do
  use ExUnit.Case, async: true

  alias UmyaSpreadsheet
  alias UmyaSpreadsheet.Table

  setup do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Create sample data in Sheet1 for table functionality
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Name")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Department")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C1", "Salary")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D1", "Start Date")

    # Row 2
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "John Doe")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", "Engineering")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C2", 75000)
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D2", "2020-01-15")

    # Row 3
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "Jane Smith")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B3", "Marketing")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C3", 65000)
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D3", "2019-03-20")

    # Row 4
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A4", "Bob Johnson")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B4", "Engineering")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C4", 80000)
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D4", "2021-06-10")

    # Row 5
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A5", "Alice Brown")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B5", "Sales")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C5", 55000)
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D5", "2022-02-01")

    {:ok, spreadsheet: spreadsheet}
  end

  test "add table to worksheet", %{spreadsheet: spreadsheet} do
    assert {:ok, :ok} = Table.add_table(
      spreadsheet,
      "Sheet1",
      "EmployeeTable",
      "Employee Data",
      "A1",
      "D5",
      ["Name", "Department", "Salary", "Start Date"],
      false
    )

    # Check if table was added
    assert {:ok, true} = Table.has_tables(spreadsheet, "Sheet1")
    assert {:ok, 1} = Table.count_tables(spreadsheet, "Sheet1")
  end

  test "get tables from worksheet", %{spreadsheet: spreadsheet} do
    # Add a table first
    assert {:ok, :ok} = Table.add_table(
      spreadsheet,
      "Sheet1",
      "EmployeeTable",
      "Employee Data",
      "A1",
      "D5",
      ["Name", "Department", "Salary", "Start Date"],
      false
    )

    # Get all tables
    {:ok, tables} = Table.get_tables(spreadsheet, "Sheet1")
    assert length(tables) == 1

    # Check table properties - tables are now maps with string keys
    table = hd(tables)
    assert table["name"] == "EmployeeTable"
    assert table["display_name"] == "Employee Data"
  end

  test "basic table operations", %{spreadsheet: spreadsheet} do
    # Test no tables initially
    assert {:ok, false} = Table.has_tables(spreadsheet, "Sheet1")
    assert {:ok, 0} = Table.count_tables(spreadsheet, "Sheet1")

    # Add a table
    assert {:ok, :ok} = Table.add_table(
      spreadsheet,
      "Sheet1",
      "EmployeeTable",
      "Employee Data",
      "A1",
      "D5",
      ["Name", "Department", "Salary", "Start Date"],
      false
    )

    # Verify table exists
    assert {:ok, true} = Table.has_tables(spreadsheet, "Sheet1")
    assert {:ok, 1} = Table.count_tables(spreadsheet, "Sheet1")

    # Remove table
    assert {:ok, :ok} = Table.remove_table(spreadsheet, "Sheet1", "EmployeeTable")

    # Verify table was removed
    assert {:ok, false} = Table.has_tables(spreadsheet, "Sheet1")
    assert {:ok, 0} = Table.count_tables(spreadsheet, "Sheet1")
  end
end
