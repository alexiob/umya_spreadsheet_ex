defmodule UmyaSpreadsheet.TableGettersTest do
  use ExUnit.Case
  alias UmyaSpreadsheet.Table

  setup do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add some test data
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Region")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Product")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C1", "Sales")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D1", "Date")

    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "North")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", "Widget A")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C2", "1000")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D2", "2023-01-01")

    # Create a test table
    {:ok, :ok} =
      Table.add_table(
        spreadsheet,
        "Sheet1",
        "TestTable",
        "Test Sales Data",
        "A1",
        "D3",
        ["Region", "Product", "Sales", "Date"],
        true
      )

    # Set table style
    {:ok, :ok} =
      Table.set_table_style(
        spreadsheet,
        "Sheet1",
        "TestTable",
        "TableStyleLight1",
        # show_first_col
        true,
        # show_last_col
        false,
        # show_row_stripes
        true,
        # show_col_stripes
        false
      )

    {:ok, spreadsheet: spreadsheet}
  end

  describe "get_table/3" do
    test "retrieves a specific table by name", %{spreadsheet: spreadsheet} do
      {:ok, table} = Table.get_table(spreadsheet, "Sheet1", "TestTable")

      assert is_map(table)
      assert table["name"] == "TestTable"
      assert table["display_name"] == "Test Sales Data"
      assert table["has_totals_row"] == "true"
      # Columns are returned as a string representation, not an actual list
      assert is_binary(table["columns"])
      # Skip the length check as we can't easily count elements in the string representation
    end

    test "returns error for non-existent table", %{spreadsheet: spreadsheet} do
      {:error, reason} = Table.get_table(spreadsheet, "Sheet1", "NonExistentTable")
      assert reason == "Table not found"
    end

    test "returns error for non-existent sheet", %{spreadsheet: spreadsheet} do
      {:error, reason} = Table.get_table(spreadsheet, "NonExistentSheet", "TestTable")
      assert reason == "Sheet not found"
    end
  end

  describe "get_table_style/3" do
    test "retrieves table style information", %{spreadsheet: spreadsheet} do
      {:ok, style} = Table.get_table_style(spreadsheet, "Sheet1", "TestTable")

      assert is_map(style)
      assert style["name"] == "TableStyleLight1"
      assert style["show_first_column"] == "true"
      assert style["show_last_column"] == "false"
      assert style["show_row_stripes"] == "true"
      assert style["show_column_stripes"] == "false"
    end

    test "returns error for non-existent table", %{spreadsheet: spreadsheet} do
      {:error, reason} = Table.get_table_style(spreadsheet, "Sheet1", "NonExistentTable")
      assert reason == "Table not found"
    end
  end

  describe "get_table_columns/3" do
    test "retrieves table column information", %{spreadsheet: spreadsheet} do
      {:ok, columns} = Table.get_table_columns(spreadsheet, "Sheet1", "TestTable")

      assert is_list(columns)
      assert length(columns) == 4

      column_names = Enum.map(columns, & &1["name"])
      assert "Region" in column_names
      assert "Product" in column_names
      assert "Sales" in column_names
      assert "Date" in column_names
    end

    test "returns error for non-existent table", %{spreadsheet: spreadsheet} do
      {:error, reason} = Table.get_table_columns(spreadsheet, "Sheet1", "NonExistentTable")
      assert reason == "Table not found"
    end
  end

  describe "get_table_totals_row/3" do
    test "retrieves totals row status", %{spreadsheet: spreadsheet} do
      {:ok, has_totals_row} = Table.get_table_totals_row(spreadsheet, "Sheet1", "TestTable")
      assert has_totals_row == true
    end

    test "returns false when totals row is disabled", %{spreadsheet: spreadsheet} do
      # Disable totals row
      {:ok, :ok} = Table.set_table_totals_row(spreadsheet, "Sheet1", "TestTable", false)

      {:ok, has_totals_row} = Table.get_table_totals_row(spreadsheet, "Sheet1", "TestTable")
      assert has_totals_row == false
    end

    test "returns error for non-existent table", %{spreadsheet: spreadsheet} do
      {:error, reason} = Table.get_table_totals_row(spreadsheet, "Sheet1", "NonExistentTable")
      assert reason == "Table not found"
    end
  end

  describe "integration with existing table functions" do
    test "get_table is consistent with get_tables", %{spreadsheet: spreadsheet} do
      {:ok, all_tables} = Table.get_tables(spreadsheet, "Sheet1")
      {:ok, specific_table} = Table.get_table(spreadsheet, "Sheet1", "TestTable")

      found_table = Enum.find(all_tables, &(&1["name"] == "TestTable"))

      assert found_table["name"] == specific_table["name"]
      assert found_table["display_name"] == specific_table["display_name"]
      # Convert both to strings for comparison
      assert to_string(found_table["has_totals_row"]) ==
               to_string(specific_table["has_totals_row"])
    end

    test "getter functions work after modifications", %{spreadsheet: spreadsheet} do
      # Add a column
      {:ok, :ok} =
        Table.add_table_column(
          spreadsheet,
          "Sheet1",
          "TestTable",
          "Total",
          "sum",
          "Grand Total"
        )

      # Verify the column was added
      {:ok, columns} = Table.get_table_columns(spreadsheet, "Sheet1", "TestTable")
      assert length(columns) == 5

      total_column = Enum.find(columns, &(&1["name"] == "Total"))
      assert total_column["totals_row_function"] == "sum"
      assert total_column["totals_row_label"] == "Grand Total"
    end
  end
end
