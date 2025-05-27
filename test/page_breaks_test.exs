defmodule UmyaSpreadsheetEx.PageBreaksTest do
  use ExUnit.Case, async: true
  doctest UmyaSpreadsheet.PageBreaks

  alias UmyaSpreadsheet.PageBreaks

  setup do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()
    sheet_name = "TestSheet"
    :ok = UmyaSpreadsheet.add_sheet(spreadsheet, sheet_name)

    %{spreadsheet: spreadsheet, sheet_name: sheet_name}
  end

  describe "single row page breaks" do
    test "add_row_page_break/4 adds a manual row page break", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      # Add a manual row page break
      assert :ok = PageBreaks.add_row_page_break(spreadsheet, sheet_name, 25)

      # Verify it exists
      assert {:ok, true} = PageBreaks.has_row_page_break(spreadsheet, sheet_name, 25)

      # Verify it's in the list of breaks
      assert {:ok, breaks} = PageBreaks.get_row_page_breaks(spreadsheet, sheet_name)
      assert {25, true} in breaks
    end

    test "add_row_page_break/4 adds an automatic row page break", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      # Add an automatic row page break
      assert :ok = PageBreaks.add_row_page_break(spreadsheet, sheet_name, 30, false)

      # Verify it exists
      assert {:ok, true} = PageBreaks.has_row_page_break(spreadsheet, sheet_name, 30)

      # Verify it's marked as automatic
      assert {:ok, breaks} = PageBreaks.get_row_page_breaks(spreadsheet, sheet_name)
      assert {30, false} in breaks
    end

    test "add_row_page_break/3 defaults to manual", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      # Add with default manual setting
      assert :ok = PageBreaks.add_row_page_break(spreadsheet, sheet_name, 20)

      # Verify it's manual
      assert {:ok, breaks} = PageBreaks.get_row_page_breaks(spreadsheet, sheet_name)
      assert {20, true} in breaks
    end

    test "remove_row_page_break/3 removes a row page break", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      # Add a row page break
      assert :ok = PageBreaks.add_row_page_break(spreadsheet, sheet_name, 15)
      assert {:ok, true} = PageBreaks.has_row_page_break(spreadsheet, sheet_name, 15)

      # Remove it
      assert :ok = PageBreaks.remove_row_page_break(spreadsheet, sheet_name, 15)
      assert {:ok, false} = PageBreaks.has_row_page_break(spreadsheet, sheet_name, 15)
    end

    test "has_row_page_break/3 returns false for non-existent break", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      assert {:ok, false} = PageBreaks.has_row_page_break(spreadsheet, sheet_name, 100)
    end

    test "page break operations with non-existent sheet return error", %{spreadsheet: spreadsheet} do
      non_existent_sheet = "NonExistentSheet"

      assert {:error, _reason} =
               PageBreaks.add_row_page_break(spreadsheet, non_existent_sheet, 10)

      assert {:error, _reason} =
               PageBreaks.remove_row_page_break(spreadsheet, non_existent_sheet, 10)

      assert {:error, _reason} =
               PageBreaks.has_row_page_break(spreadsheet, non_existent_sheet, 10)

      assert {:error, _reason} = PageBreaks.get_row_page_breaks(spreadsheet, non_existent_sheet)
    end
  end

  describe "single column page breaks" do
    test "add_column_page_break/4 adds a manual column page break", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      # Add a manual column page break
      assert :ok = PageBreaks.add_column_page_break(spreadsheet, sheet_name, 10)

      # Verify it exists
      assert {:ok, true} = PageBreaks.has_column_page_break(spreadsheet, sheet_name, 10)

      # Verify it's in the list of breaks
      assert {:ok, breaks} = PageBreaks.get_column_page_breaks(spreadsheet, sheet_name)
      assert {10, true} in breaks
    end

    test "add_column_page_break/4 adds an automatic column page break", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      # Add an automatic column page break
      assert :ok = PageBreaks.add_column_page_break(spreadsheet, sheet_name, 8, false)

      # Verify it exists
      assert {:ok, true} = PageBreaks.has_column_page_break(spreadsheet, sheet_name, 8)

      # Verify it's marked as automatic
      assert {:ok, breaks} = PageBreaks.get_column_page_breaks(spreadsheet, sheet_name)
      assert {8, false} in breaks
    end

    test "remove_column_page_break/3 removes a column page break", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      # Add a column page break
      assert :ok = PageBreaks.add_column_page_break(spreadsheet, sheet_name, 12)
      assert {:ok, true} = PageBreaks.has_column_page_break(spreadsheet, sheet_name, 12)

      # Remove it
      assert :ok = PageBreaks.remove_column_page_break(spreadsheet, sheet_name, 12)
      assert {:ok, false} = PageBreaks.has_column_page_break(spreadsheet, sheet_name, 12)
    end

    test "has_column_page_break/3 returns false for non-existent break", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      assert {:ok, false} = PageBreaks.has_column_page_break(spreadsheet, sheet_name, 50)
    end
  end

  describe "clearing page breaks" do
    test "clear_row_page_breaks/2 removes all row page breaks", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      # Add multiple row page breaks
      assert :ok = PageBreaks.add_row_page_break(spreadsheet, sheet_name, 10)
      assert :ok = PageBreaks.add_row_page_break(spreadsheet, sheet_name, 20)
      assert :ok = PageBreaks.add_row_page_break(spreadsheet, sheet_name, 30)

      # Verify they exist
      assert {:ok, breaks} = PageBreaks.get_row_page_breaks(spreadsheet, sheet_name)
      assert length(breaks) == 3

      # Clear all row breaks
      assert :ok = PageBreaks.clear_row_page_breaks(spreadsheet, sheet_name)

      # Verify they're all gone
      assert {:ok, breaks} = PageBreaks.get_row_page_breaks(spreadsheet, sheet_name)
      assert breaks == []
    end

    test "clear_column_page_breaks/2 removes all column page breaks", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      # Add multiple column page breaks
      assert :ok = PageBreaks.add_column_page_break(spreadsheet, sheet_name, 5)
      assert :ok = PageBreaks.add_column_page_break(spreadsheet, sheet_name, 10)
      assert :ok = PageBreaks.add_column_page_break(spreadsheet, sheet_name, 15)

      # Verify they exist
      assert {:ok, breaks} = PageBreaks.get_column_page_breaks(spreadsheet, sheet_name)
      assert length(breaks) == 3

      # Clear all column breaks
      assert :ok = PageBreaks.clear_column_page_breaks(spreadsheet, sheet_name)

      # Verify they're all gone
      assert {:ok, breaks} = PageBreaks.get_column_page_breaks(spreadsheet, sheet_name)
      assert breaks == []
    end
  end

  describe "counting page breaks" do
    test "count_row_page_breaks/2 returns correct count", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      # Initially no breaks
      assert {:ok, 0} = PageBreaks.count_row_page_breaks(spreadsheet, sheet_name)

      # Add some breaks
      assert :ok = PageBreaks.add_row_page_break(spreadsheet, sheet_name, 10)
      assert :ok = PageBreaks.add_row_page_break(spreadsheet, sheet_name, 20)
      assert :ok = PageBreaks.add_row_page_break(spreadsheet, sheet_name, 30)

      # Verify count
      assert {:ok, 3} = PageBreaks.count_row_page_breaks(spreadsheet, sheet_name)

      # Remove one
      assert :ok = PageBreaks.remove_row_page_break(spreadsheet, sheet_name, 20)
      assert {:ok, 2} = PageBreaks.count_row_page_breaks(spreadsheet, sheet_name)
    end

    test "count_column_page_breaks/2 returns correct count", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      # Initially no breaks
      assert {:ok, 0} = PageBreaks.count_column_page_breaks(spreadsheet, sheet_name)

      # Add some breaks
      assert :ok = PageBreaks.add_column_page_break(spreadsheet, sheet_name, 5)
      assert :ok = PageBreaks.add_column_page_break(spreadsheet, sheet_name, 10)

      # Verify count
      assert {:ok, 2} = PageBreaks.count_column_page_breaks(spreadsheet, sheet_name)

      # Clear all
      assert :ok = PageBreaks.clear_column_page_breaks(spreadsheet, sheet_name)
      assert {:ok, 0} = PageBreaks.count_column_page_breaks(spreadsheet, sheet_name)
    end
  end

  describe "bulk operations" do
    test "add_row_page_breaks/3 adds multiple row page breaks", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      row_numbers = [10, 20, 30, 40, 50]

      # Add all breaks in one operation
      assert :ok = PageBreaks.add_row_page_breaks(spreadsheet, sheet_name, row_numbers)

      # Verify all breaks were added
      assert {:ok, breaks} = PageBreaks.get_row_page_breaks(spreadsheet, sheet_name)
      break_rows = Enum.map(breaks, fn {row, _manual} -> row end)

      Enum.each(row_numbers, fn row ->
        assert row in break_rows
        assert {:ok, true} = PageBreaks.has_row_page_break(spreadsheet, sheet_name, row)
      end)

      # Verify count
      assert {:ok, 5} = PageBreaks.count_row_page_breaks(spreadsheet, sheet_name)
    end

    test "add_column_page_breaks/3 adds multiple column page breaks", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      column_numbers = [5, 10, 15, 20]

      # Add all breaks in one operation
      assert :ok = PageBreaks.add_column_page_breaks(spreadsheet, sheet_name, column_numbers)

      # Verify all breaks were added
      assert {:ok, breaks} = PageBreaks.get_column_page_breaks(spreadsheet, sheet_name)
      break_columns = Enum.map(breaks, fn {col, _manual} -> col end)

      Enum.each(column_numbers, fn col ->
        assert col in break_columns
        assert {:ok, true} = PageBreaks.has_column_page_break(spreadsheet, sheet_name, col)
      end)

      # Verify count
      assert {:ok, 4} = PageBreaks.count_column_page_breaks(spreadsheet, sheet_name)
    end

    test "remove_row_page_breaks/3 removes multiple row page breaks", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      # Add multiple breaks
      all_rows = [10, 20, 30, 40, 50]
      assert :ok = PageBreaks.add_row_page_breaks(spreadsheet, sheet_name, all_rows)

      # Remove some breaks
      remove_rows = [20, 40]
      assert :ok = PageBreaks.remove_row_page_breaks(spreadsheet, sheet_name, remove_rows)

      # Verify specific breaks were removed
      assert {:ok, false} = PageBreaks.has_row_page_break(spreadsheet, sheet_name, 20)
      assert {:ok, false} = PageBreaks.has_row_page_break(spreadsheet, sheet_name, 40)

      # Verify others remain
      assert {:ok, true} = PageBreaks.has_row_page_break(spreadsheet, sheet_name, 10)
      assert {:ok, true} = PageBreaks.has_row_page_break(spreadsheet, sheet_name, 30)
      assert {:ok, true} = PageBreaks.has_row_page_break(spreadsheet, sheet_name, 50)

      # Verify count
      assert {:ok, 3} = PageBreaks.count_row_page_breaks(spreadsheet, sheet_name)
    end

    test "remove_column_page_breaks/3 removes multiple column page breaks", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      # Add multiple breaks
      all_columns = [5, 10, 15, 20, 25]
      assert :ok = PageBreaks.add_column_page_breaks(spreadsheet, sheet_name, all_columns)

      # Remove some breaks
      remove_columns = [10, 20]
      assert :ok = PageBreaks.remove_column_page_breaks(spreadsheet, sheet_name, remove_columns)

      # Verify specific breaks were removed
      assert {:ok, false} = PageBreaks.has_column_page_break(spreadsheet, sheet_name, 10)
      assert {:ok, false} = PageBreaks.has_column_page_break(spreadsheet, sheet_name, 20)

      # Verify others remain
      assert {:ok, true} = PageBreaks.has_column_page_break(spreadsheet, sheet_name, 5)
      assert {:ok, true} = PageBreaks.has_column_page_break(spreadsheet, sheet_name, 15)
      assert {:ok, true} = PageBreaks.has_column_page_break(spreadsheet, sheet_name, 25)

      # Verify count
      assert {:ok, 3} = PageBreaks.count_column_page_breaks(spreadsheet, sheet_name)
    end

    test "bulk operations with empty lists work correctly", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      # Adding empty lists should succeed but do nothing
      assert :ok = PageBreaks.add_row_page_breaks(spreadsheet, sheet_name, [])
      assert :ok = PageBreaks.add_column_page_breaks(spreadsheet, sheet_name, [])
      assert :ok = PageBreaks.remove_row_page_breaks(spreadsheet, sheet_name, [])
      assert :ok = PageBreaks.remove_column_page_breaks(spreadsheet, sheet_name, [])

      # Verify no breaks exist
      assert {:ok, 0} = PageBreaks.count_row_page_breaks(spreadsheet, sheet_name)
      assert {:ok, 0} = PageBreaks.count_column_page_breaks(spreadsheet, sheet_name)
    end

    test "bulk operations fail if one operation fails", %{spreadsheet: spreadsheet} do
      non_existent_sheet = "NonExistentSheet"

      # These should all fail because the sheet doesn't exist
      assert {:error, _reason} =
               PageBreaks.add_row_page_breaks(spreadsheet, non_existent_sheet, [10, 20])

      assert {:error, _reason} =
               PageBreaks.add_column_page_breaks(spreadsheet, non_existent_sheet, [5, 10])

      assert {:error, _reason} =
               PageBreaks.remove_row_page_breaks(spreadsheet, non_existent_sheet, [10])

      assert {:error, _reason} =
               PageBreaks.remove_column_page_breaks(spreadsheet, non_existent_sheet, [5])
    end
  end

  describe "convenience functions" do
    test "clear_all_page_breaks/2 clears both row and column breaks", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      # Add both row and column breaks
      assert :ok = PageBreaks.add_row_page_breaks(spreadsheet, sheet_name, [10, 20, 30])
      assert :ok = PageBreaks.add_column_page_breaks(spreadsheet, sheet_name, [5, 10, 15])

      # Verify they exist
      assert {:ok, row_count} = PageBreaks.count_row_page_breaks(spreadsheet, sheet_name)
      assert {:ok, col_count} = PageBreaks.count_column_page_breaks(spreadsheet, sheet_name)
      assert row_count > 0
      assert col_count > 0

      # Clear all breaks
      assert :ok = PageBreaks.clear_all_page_breaks(spreadsheet, sheet_name)

      # Verify all are gone
      assert {:ok, 0} = PageBreaks.count_row_page_breaks(spreadsheet, sheet_name)
      assert {:ok, 0} = PageBreaks.count_column_page_breaks(spreadsheet, sheet_name)
    end

    test "get_all_page_breaks/2 returns both row and column breaks", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      # Add some breaks
      row_numbers = [15, 30, 45]
      column_numbers = [8, 16]

      assert :ok = PageBreaks.add_row_page_breaks(spreadsheet, sheet_name, row_numbers)
      assert :ok = PageBreaks.add_column_page_breaks(spreadsheet, sheet_name, column_numbers)

      # Get all breaks at once
      assert {:ok, %{row_breaks: row_breaks, column_breaks: column_breaks}} =
               PageBreaks.get_all_page_breaks(spreadsheet, sheet_name)

      # Verify structure and content
      assert is_list(row_breaks)
      assert is_list(column_breaks)
      assert length(row_breaks) == 3
      assert length(column_breaks) == 2

      # Verify row breaks
      row_positions = Enum.map(row_breaks, fn {row, _manual} -> row end)
      Enum.each(row_numbers, fn row -> assert row in row_positions end)

      # Verify column breaks
      col_positions = Enum.map(column_breaks, fn {col, _manual} -> col end)
      Enum.each(column_numbers, fn col -> assert col in col_positions end)
    end

    test "get_all_page_breaks/2 returns empty lists when no breaks exist", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      assert {:ok, %{row_breaks: [], column_breaks: []}} =
               PageBreaks.get_all_page_breaks(spreadsheet, sheet_name)
    end

    test "get_all_page_breaks/2 fails with non-existent sheet", %{spreadsheet: spreadsheet} do
      assert {:error, _reason} = PageBreaks.get_all_page_breaks(spreadsheet, "NonExistentSheet")
    end
  end

  describe "manual vs automatic page breaks" do
    test "can distinguish between manual and automatic breaks", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      # Add manual and automatic breaks
      # manual
      assert :ok = PageBreaks.add_row_page_break(spreadsheet, sheet_name, 10, true)
      # automatic
      assert :ok = PageBreaks.add_row_page_break(spreadsheet, sheet_name, 20, false)
      # manual
      assert :ok = PageBreaks.add_column_page_break(spreadsheet, sheet_name, 5, true)
      # automatic
      assert :ok = PageBreaks.add_column_page_break(spreadsheet, sheet_name, 10, false)

      # Get breaks and verify manual/automatic flags
      assert {:ok, row_breaks} = PageBreaks.get_row_page_breaks(spreadsheet, sheet_name)
      assert {:ok, col_breaks} = PageBreaks.get_column_page_breaks(spreadsheet, sheet_name)

      # Check row breaks
      # manual
      assert {10, true} in row_breaks
      # automatic
      assert {20, false} in row_breaks

      # Check column breaks
      # manual
      assert {5, true} in col_breaks
      # automatic
      assert {10, false} in col_breaks
    end

    test "bulk operations create manual breaks by default", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      # Add breaks using bulk operations
      assert :ok = PageBreaks.add_row_page_breaks(spreadsheet, sheet_name, [25, 50])
      assert :ok = PageBreaks.add_column_page_breaks(spreadsheet, sheet_name, [12, 24])

      # Verify all are manual
      assert {:ok, row_breaks} = PageBreaks.get_row_page_breaks(spreadsheet, sheet_name)
      assert {:ok, col_breaks} = PageBreaks.get_column_page_breaks(spreadsheet, sheet_name)

      Enum.each(row_breaks, fn {_row, manual} -> assert manual == true end)
      Enum.each(col_breaks, fn {_col, manual} -> assert manual == true end)
    end
  end

  describe "edge cases and error handling" do
    test "can add multiple breaks at same position (should replace)", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      # Add manual break
      assert :ok = PageBreaks.add_row_page_break(spreadsheet, sheet_name, 15, true)
      assert {:ok, breaks} = PageBreaks.get_row_page_breaks(spreadsheet, sheet_name)
      assert {15, true} in breaks

      # Add automatic break at same position (this might replace or be ignored)
      assert :ok = PageBreaks.add_row_page_break(spreadsheet, sheet_name, 15, false)
      assert {:ok, breaks} = PageBreaks.get_row_page_breaks(spreadsheet, sheet_name)

      # Should still have a break at position 15
      break_positions = Enum.map(breaks, fn {row, _} -> row end)
      assert 15 in break_positions
    end

    test "removing non-existent break succeeds", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      # Should not error when removing a break that doesn't exist
      assert :ok = PageBreaks.remove_row_page_break(spreadsheet, sheet_name, 999)
      assert :ok = PageBreaks.remove_column_page_break(spreadsheet, sheet_name, 999)
    end

    test "can handle large numbers of page breaks", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      # Add many page breaks
      # [10, 20, 30, ..., 100]
      large_row_list = Enum.to_list(10..100//10)
      # [5, 10, 15, ..., 50]
      large_col_list = Enum.to_list(5..50//5)

      assert :ok = PageBreaks.add_row_page_breaks(spreadsheet, sheet_name, large_row_list)
      assert :ok = PageBreaks.add_column_page_breaks(spreadsheet, sheet_name, large_col_list)

      # Verify counts
      assert {:ok, 10} = PageBreaks.count_row_page_breaks(spreadsheet, sheet_name)
      assert {:ok, 10} = PageBreaks.count_column_page_breaks(spreadsheet, sheet_name)

      # Verify we can get all breaks
      assert {:ok, %{row_breaks: row_breaks, column_breaks: col_breaks}} =
               PageBreaks.get_all_page_breaks(spreadsheet, sheet_name)

      assert length(row_breaks) == 10
      assert length(col_breaks) == 10
    end
  end

  describe "integration with real spreadsheet operations" do
    test "page breaks persist through spreadsheet operations", %{
      spreadsheet: spreadsheet,
      sheet_name: sheet_name
    } do
      # Add some data to the sheet
      :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, "A1", "Header")
      :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, "A10", "Section 1")
      :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, "A20", "Section 2")

      # Add page breaks
      assert :ok = PageBreaks.add_row_page_breaks(spreadsheet, sheet_name, [10, 20])
      assert :ok = PageBreaks.add_column_page_breaks(spreadsheet, sheet_name, [5])

      # Add more data
      :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, "B1", "Data")

      # Verify page breaks still exist
      assert {:ok, %{row_breaks: row_breaks, column_breaks: col_breaks}} =
               PageBreaks.get_all_page_breaks(spreadsheet, sheet_name)

      row_positions = Enum.map(row_breaks, fn {row, _} -> row end)
      col_positions = Enum.map(col_breaks, fn {col, _} -> col end)

      assert 10 in row_positions
      assert 20 in row_positions
      assert 5 in col_positions
    end

    test "can work with multiple sheets independently", %{spreadsheet: spreadsheet} do
      # Create multiple sheets with unique names
      sheet1 = "TestSheet1"
      sheet2 = "TestSheet2"
      :ok = UmyaSpreadsheet.add_sheet(spreadsheet, sheet1)
      :ok = UmyaSpreadsheet.add_sheet(spreadsheet, sheet2)

      # Add different page breaks to each sheet
      assert :ok = PageBreaks.add_row_page_breaks(spreadsheet, sheet1, [10, 20])
      assert :ok = PageBreaks.add_row_page_breaks(spreadsheet, sheet2, [15, 25, 35])

      assert :ok = PageBreaks.add_column_page_breaks(spreadsheet, sheet1, [5])
      assert :ok = PageBreaks.add_column_page_breaks(spreadsheet, sheet2, [8, 16])

      # Verify each sheet has its own breaks
      assert {:ok, %{row_breaks: sheet1_rows, column_breaks: sheet1_cols}} =
               PageBreaks.get_all_page_breaks(spreadsheet, sheet1)

      assert {:ok, %{row_breaks: sheet2_rows, column_breaks: sheet2_cols}} =
               PageBreaks.get_all_page_breaks(spreadsheet, sheet2)

      # Verify counts
      assert length(sheet1_rows) == 2
      assert length(sheet1_cols) == 1
      assert length(sheet2_rows) == 3
      assert length(sheet2_cols) == 2

      # Clear breaks from one sheet shouldn't affect the other
      assert :ok = PageBreaks.clear_all_page_breaks(spreadsheet, sheet1)

      assert {:ok, %{row_breaks: [], column_breaks: []}} =
               PageBreaks.get_all_page_breaks(spreadsheet, sheet1)

      assert {:ok, %{row_breaks: sheet2_rows_after, column_breaks: sheet2_cols_after}} =
               PageBreaks.get_all_page_breaks(spreadsheet, sheet2)

      # Sheet2 should be unchanged
      assert length(sheet2_rows_after) == 3
      assert length(sheet2_cols_after) == 2
    end
  end
end
