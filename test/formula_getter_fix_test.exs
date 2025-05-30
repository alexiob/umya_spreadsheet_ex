defmodule UmyaSpreadsheetTest.FormulaGetterFixTest do
  use ExUnit.Case, async: true

  alias UmyaSpreadsheet

  describe "formula getter functions" do
    setup do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Set a formula in a cell for testing
      :ok = UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "A1", "=SUM(B1:B5)")

      %{spreadsheet: spreadsheet}
    end

    test "is_formula function works without ArgumentError", %{spreadsheet: spreadsheet} do
      # Test is_formula - this was one of the functions we fixed
      result = UmyaSpreadsheet.is_formula(spreadsheet, "Sheet1", "A1")
      assert is_boolean(result)
      assert result == true

      # Test with a non-formula cell
      result2 = UmyaSpreadsheet.is_formula(spreadsheet, "Sheet1", "B1")
      assert is_boolean(result2)
      assert result2 == false
    end

    test "get_formula function works without ArgumentError", %{spreadsheet: spreadsheet} do
      # Test get_formula - this was one of the functions we fixed
      result = UmyaSpreadsheet.get_formula(spreadsheet, "Sheet1", "A1")
      assert is_binary(result)
      # Formula includes the equals sign
      assert result == "=SUM(B1:B5)"
    end

    test "get_text function works without ArgumentError", %{spreadsheet: spreadsheet} do
      # Test get_text - this was one of the functions we fixed
      result = UmyaSpreadsheet.get_text(spreadsheet, "Sheet1", "A1")
      assert is_binary(result)
    end

    test "get_formula_obj function works without ArgumentError", %{spreadsheet: spreadsheet} do
      # Test get_formula_obj - this was one of the functions we fixed
      # The function returns a 4-tuple: {text, type, shared_index, reference}
      {formula, type, _shared_index, _reference} =
        UmyaSpreadsheet.get_formula_obj(spreadsheet, "Sheet1", "A1")

      assert is_binary(formula)
      assert is_binary(type)
      # Formula includes the equals sign
      assert formula == "=SUM(B1:B5)"
      assert type == "Normal"
    end

    test "get_formula_shared_index function works without ArgumentError", %{
      spreadsheet: spreadsheet
    } do
      # Test get_formula_shared_index - this was one of the functions we fixed
      result = UmyaSpreadsheet.get_formula_shared_index(spreadsheet, "Sheet1", "A1")
      # This might return nil for a non-shared formula
      assert is_nil(result) or is_integer(result)
    end
  end
end
