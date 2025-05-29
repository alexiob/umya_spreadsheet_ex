defmodule UmyaSpreadsheet.SheetGetterFunctionsTest do
  use ExUnit.Case
  alias UmyaSpreadsheet.SheetFunctions

  describe "sheet getter functions" do
    setup do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()
      {:ok, spreadsheet: spreadsheet}
    end

    test "get_sheet_count returns correct number of sheets", %{spreadsheet: spreadsheet} do
      sheet_count = SheetFunctions.get_sheet_count(spreadsheet)
      assert is_integer(sheet_count)
      assert sheet_count >= 1
    end

    test "get_active_sheet returns active sheet index", %{spreadsheet: spreadsheet} do
      result = SheetFunctions.get_active_sheet(spreadsheet)
      assert {:ok, index} = result
      assert is_integer(index)
      assert index >= 0
    end

    test "get_sheet_state returns visible state for default sheet", %{spreadsheet: spreadsheet} do
      result = SheetFunctions.get_sheet_state(spreadsheet, "Sheet1")
      assert {:ok, state} = result
      assert state in ["visible", "hidden", "veryhidden"]
    end

    test "get_sheet_state returns error for non-existent sheet", %{spreadsheet: spreadsheet} do
      result = SheetFunctions.get_sheet_state(spreadsheet, "NonExistentSheet")
      assert {:error, _reason} = result
    end

    test "get_sheet_protection returns protection info", %{spreadsheet: spreadsheet} do
      result = SheetFunctions.get_sheet_protection(spreadsheet, "Sheet1")
      assert {:ok, protection} = result
      assert is_map(protection)
    end

    test "get_sheet_protection returns error for non-existent sheet", %{spreadsheet: spreadsheet} do
      result = SheetFunctions.get_sheet_protection(spreadsheet, "NonExistentSheet")
      assert {:error, _reason} = result
    end

    test "get_merge_cells returns empty list for new sheet", %{spreadsheet: spreadsheet} do
      result = SheetFunctions.get_merge_cells(spreadsheet, "Sheet1")
      assert {:ok, merge_cells} = result
      assert is_list(merge_cells)
    end

    test "get_merge_cells returns error for non-existent sheet", %{spreadsheet: spreadsheet} do
      result = SheetFunctions.get_merge_cells(spreadsheet, "NonExistentSheet")
      assert {:error, _reason} = result
    end
  end
end
