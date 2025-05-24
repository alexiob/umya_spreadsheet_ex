defmodule UmyaSpreadsheetTest.FormulaTest do
  use ExUnit.Case, async: true

  alias UmyaSpreadsheet

  setup do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()
    # Set some values to use in formulas
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", 10)
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", 20)
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B3", 30)
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B4", 40)
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B5", 50)
    %{spreadsheet: spreadsheet}
  end

  test "set_formula sets a regular formula in a cell", %{spreadsheet: spreadsheet} do
    # Set a simple formula
    assert :ok = UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "C1", "SUM(B1:B5)")

    # Can't test the result value as the formulas aren't calculated by the library
    # but we can test that the function returns :ok
  end

  test "set_array_formula sets an array formula for a range", %{spreadsheet: spreadsheet} do
    # Set an array formula that should fill multiple cells
    assert :ok = UmyaSpreadsheet.set_array_formula(spreadsheet, "Sheet1", "D1:D3", "B1:B3")

    # Can't test the result values as the formulas aren't calculated by the library
    # but we can test that the function returns :ok
  end

  test "create_named_range creates a named range", %{spreadsheet: spreadsheet} do
    # Create a named range
    assert :ok = UmyaSpreadsheet.create_named_range(spreadsheet, "DataRange", "Sheet1", "B1:B5")

    # Currently can't test retrieving the named range as the library doesn't support it yet
  end

  test "create_defined_name creates a defined name with a formula", %{spreadsheet: spreadsheet} do
    # Create a defined name for a constant
    assert :ok = UmyaSpreadsheet.create_defined_name(spreadsheet, "TaxRate", "0.15")

    # Create a defined name scoped to a specific sheet
    assert :ok = UmyaSpreadsheet.create_defined_name(spreadsheet, "Subtotal", "SUM(B1:B3)", "Sheet1")

    # Currently can't test retrieving the defined name as the library doesn't support it yet
  end

  test "get_defined_names retrieves defined names", %{spreadsheet: spreadsheet} do
    # Create some defined names
    assert :ok = UmyaSpreadsheet.create_named_range(spreadsheet, "DataRange", "Sheet1", "B1:B5")
    assert :ok = UmyaSpreadsheet.create_defined_name(spreadsheet, "TaxRate", "0.15")
    assert :ok = UmyaSpreadsheet.create_defined_name(spreadsheet, "Subtotal", "SUM(B1:B3)", "Sheet1")

    # Get the defined names
    defined_names = UmyaSpreadsheet.get_defined_names(spreadsheet)
    assert is_list(defined_names)

    # Check that we have the expected number of defined names
    assert length(defined_names) == 3

    # Check that our named range and sheet-scoped defined name have proper names
    assert Enum.any?(defined_names, fn {name, _} -> name == "DataRange" end)
    assert Enum.any?(defined_names, fn {name, _} -> name == "Subtotal" end)

    # Check that global defined name exists (but note: has empty name due to API limitation)
    assert Enum.any?(defined_names, fn {_, formula} -> String.contains?(formula, "0.15") end)

    # Find the specific items and check their formulas
    data_range = Enum.find(defined_names, fn {name, _} -> name == "DataRange" end)
    subtotal = Enum.find(defined_names, fn {name, _} -> name == "Subtotal" end)
    # Global defined names have empty names due to umya-spreadsheet 2.3.0 API limitation
    tax_rate = Enum.find(defined_names, fn {name, formula} ->
      name == "" && String.contains?(formula, "0.15")
    end)

    assert data_range != nil
    assert tax_rate != nil
    assert subtotal != nil

    {_, data_range_formula} = data_range
    {_, tax_rate_formula} = tax_rate
    {_, subtotal_formula} = subtotal

    # Check formula contents (may vary slightly in format from what we set)
    assert String.contains?(data_range_formula, "B1:B5")
    assert String.contains?(tax_rate_formula, "0.15")
    assert String.contains?(subtotal_formula, "SUM(B1:B3)")
  end
end
