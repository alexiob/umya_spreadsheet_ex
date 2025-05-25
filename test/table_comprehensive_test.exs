defmodule UmyaSpreadsheet.TableComprehensiveTest do
  use ExUnit.Case, async: true
  alias UmyaSpreadsheet.Table

  describe "Table functionality with tuple returns" do
    setup do
      # Create a new spreadsheet for each test
      {:ok, spreadsheet} = UmyaSpreadsheet.new()
      %{spreadsheet: spreadsheet}
    end

    test "add_table returns proper tuple", %{spreadsheet: spreadsheet} do
      # Add a table and verify it returns the expected tuple
      result = Table.add_table(
        spreadsheet,
        "Sheet1",
        "TestTable",
        "Test Table",
        "A1",
        "C5",
        ["Name", "Age", "City"],
        true
      )

      assert {:ok, :ok} = result
    end

    test "get_tables returns proper tuple", %{spreadsheet: spreadsheet} do
      # First add a table
      {:ok, :ok} = Table.add_table(
        spreadsheet,
        "Sheet1",
        "TestTable",
        "Test Table",
        "A1",
        "C5",
        ["Name", "Age", "City"],
        true
      )

      # Get tables and verify the return format
      result = Table.get_tables(spreadsheet, "Sheet1")
      
      assert {:ok, tables} = result
      assert is_list(tables)
      assert length(tables) == 1
      
      [table | _] = tables
      assert Map.has_key?(table, "name")
      assert Map.has_key?(table, "display_name")
      assert Map.has_key?(table, "columns")
      assert table["name"] == "TestTable"
      assert table["display_name"] == "Test Table"
    end

    test "has_tables returns proper tuple", %{spreadsheet: spreadsheet} do
      # Check empty sheet
      result = Table.has_tables(spreadsheet, "Sheet1")
      assert {:ok, false} = result

      # Add a table
      {:ok, :ok} = Table.add_table(
        spreadsheet,
        "Sheet1",
        "TestTable",
        "Test Table",
        "A1",
        "C5",
        ["Name", "Age", "City"]
      )

      # Check again
      result = Table.has_tables(spreadsheet, "Sheet1")
      assert {:ok, true} = result
    end

    test "count_tables returns proper tuple", %{spreadsheet: spreadsheet} do
      # Check empty sheet
      result = Table.count_tables(spreadsheet, "Sheet1")
      assert {:ok, 0} = result

      # Add a table
      {:ok, :ok} = Table.add_table(
        spreadsheet,
        "Sheet1",
        "TestTable1",
        "Test Table 1",
        "A1",
        "C5",
        ["Name", "Age", "City"]
      )

      # Check count
      result = Table.count_tables(spreadsheet, "Sheet1")
      assert {:ok, 1} = result

      # Add another table
      {:ok, :ok} = Table.add_table(
        spreadsheet,
        "Sheet1",
        "TestTable2",
        "Test Table 2",
        "E1",
        "G5",
        ["Product", "Price", "Stock"]
      )

      # Check count again
      result = Table.count_tables(spreadsheet, "Sheet1")
      assert {:ok, 2} = result
    end

    test "set_table_style returns proper tuple", %{spreadsheet: spreadsheet} do
      # Add a table first
      {:ok, :ok} = Table.add_table(
        spreadsheet,
        "Sheet1",
        "TestTable",
        "Test Table",
        "A1",
        "C5",
        ["Name", "Age", "City"]
      )

      # Set table style
      result = Table.set_table_style(
        spreadsheet,
        "Sheet1",
        "TestTable",
        "TableStyleMedium2",
        true,
        false,
        true,
        false
      )

      assert {:ok, :ok} = result
    end

    test "remove_table_style returns proper tuple", %{spreadsheet: spreadsheet} do
      # Add a table first
      {:ok, :ok} = Table.add_table(
        spreadsheet,
        "Sheet1",
        "TestTable",
        "Test Table",
        "A1",
        "C5",
        ["Name", "Age", "City"]
      )

      # Remove table style
      result = Table.remove_table_style(spreadsheet, "Sheet1", "TestTable")
      assert {:ok, :ok} = result
    end

    test "add_table_column returns proper tuple", %{spreadsheet: spreadsheet} do
      # Add a table first
      {:ok, :ok} = Table.add_table(
        spreadsheet,
        "Sheet1",
        "TestTable",
        "Test Table",
        "A1",
        "C5",
        ["Name", "Age", "City"]
      )

      # Add a column
      result = Table.add_table_column(
        spreadsheet,
        "Sheet1",
        "TestTable",
        "Email",
        "count",
        "Total Emails"
      )

      assert {:ok, :ok} = result
    end

    test "modify_table_column returns proper tuple", %{spreadsheet: spreadsheet} do
      # Add a table first
      {:ok, :ok} = Table.add_table(
        spreadsheet,
        "Sheet1",
        "TestTable",
        "Test Table",
        "A1",
        "C5",
        ["Name", "Age", "City"]
      )

      # Modify a column
      result = Table.modify_table_column(
        spreadsheet,
        "Sheet1",
        "TestTable",
        "Age",
        "Years",
        "average",
        "Avg Age"
      )

      assert {:ok, :ok} = result
    end

    test "set_table_totals_row returns proper tuple", %{spreadsheet: spreadsheet} do
      # Add a table first
      {:ok, :ok} = Table.add_table(
        spreadsheet,
        "Sheet1",
        "TestTable",
        "Test Table",
        "A1",
        "C5",
        ["Name", "Age", "City"]
      )

      # Enable totals row
      result = Table.set_table_totals_row(spreadsheet, "Sheet1", "TestTable", true)
      assert {:ok, :ok} = result

      # Disable totals row
      result = Table.set_table_totals_row(spreadsheet, "Sheet1", "TestTable", false)
      assert {:ok, :ok} = result
    end

    test "remove_table returns proper tuple", %{spreadsheet: spreadsheet} do
      # Add a table first
      {:ok, :ok} = Table.add_table(
        spreadsheet,
        "Sheet1",
        "TestTable",
        "Test Table",
        "A1",
        "C5",
        ["Name", "Age", "City"]
      )

      # Verify table exists
      {:ok, 1} = Table.count_tables(spreadsheet, "Sheet1")

      # Remove the table
      result = Table.remove_table(spreadsheet, "Sheet1", "TestTable")
      assert {:ok, :ok} = result

      # Verify table is gone
      {:ok, 0} = Table.count_tables(spreadsheet, "Sheet1")
    end

    test "error cases return proper error tuples", %{spreadsheet: spreadsheet} do
      # Test with non-existent sheet
      result = Table.get_tables(spreadsheet, "NonExistentSheet")
      assert {:error, "Sheet not found"} = result

      result = Table.has_tables(spreadsheet, "NonExistentSheet")
      assert {:error, "Sheet not found"} = result

      result = Table.count_tables(spreadsheet, "NonExistentSheet")
      assert {:error, "Sheet not found"} = result

      # Test with non-existent table
      result = Table.remove_table(spreadsheet, "Sheet1", "NonExistentTable")
      assert {:error, "Table not found"} = result

      result = Table.set_table_style(
        spreadsheet,
        "Sheet1",
        "NonExistentTable",
        "TableStyleMedium2",
        true,
        false,
        true,
        false
      )
      assert {:error, "Table not found"} = result
    end

    test "complete table workflow", %{spreadsheet: spreadsheet} do
      # 1. Verify no tables initially
      {:ok, false} = Table.has_tables(spreadsheet, "Sheet1")
      {:ok, 0} = Table.count_tables(spreadsheet, "Sheet1")

      # 2. Add a table
      {:ok, :ok} = Table.add_table(
        spreadsheet,
        "Sheet1",
        "SalesTable",
        "Sales Data",
        "A1",
        "E10",
        ["Region", "Product", "Sales", "Date", "Rep"],
        true
      )

      # 3. Verify table was added
      {:ok, true} = Table.has_tables(spreadsheet, "Sheet1")
      {:ok, 1} = Table.count_tables(spreadsheet, "Sheet1")

      # 4. Get table details
      {:ok, [table]} = Table.get_tables(spreadsheet, "Sheet1")
      assert table["name"] == "SalesTable"
      assert table["display_name"] == "Sales Data"
      assert length(table["columns"]) == 5

      # 5. Style the table
      {:ok, :ok} = Table.set_table_style(
        spreadsheet,
        "Sheet1",
        "SalesTable",
        "TableStyleMedium9",
        true,
        true,
        true,
        false
      )

      # 6. Add a column
      {:ok, :ok} = Table.add_table_column(
        spreadsheet,
        "Sheet1",
        "SalesTable",
        "Commission",
        "sum",
        "Total Commission"
      )

      # 7. Modify a column
      {:ok, :ok} = Table.modify_table_column(
        spreadsheet,
        "Sheet1",
        "SalesTable",
        "Sales",
        "Revenue",
        "sum",
        "Total Revenue"
      )

      # 8. Toggle totals row
      {:ok, :ok} = Table.set_table_totals_row(spreadsheet, "Sheet1", "SalesTable", false)
      {:ok, :ok} = Table.set_table_totals_row(spreadsheet, "Sheet1", "SalesTable", true)

      # 9. Remove styling
      {:ok, :ok} = Table.remove_table_style(spreadsheet, "Sheet1", "SalesTable")

      # 10. Finally remove the table
      {:ok, :ok} = Table.remove_table(spreadsheet, "Sheet1", "SalesTable")
      {:ok, false} = Table.has_tables(spreadsheet, "Sheet1")
      {:ok, 0} = Table.count_tables(spreadsheet, "Sheet1")
    end
  end
end
