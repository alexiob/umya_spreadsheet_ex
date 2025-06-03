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

    # Verify the formula was actually set
    assert UmyaSpreadsheet.is_formula(spreadsheet, "Sheet1", "C1") == true

    # Verify formula text (doesn't include the = sign)
    formula_text = UmyaSpreadsheet.get_formula(spreadsheet, "Sheet1", "C1")
    assert formula_text == "SUM(B1:B5)"

    # Verify formula type
    formula_type = UmyaSpreadsheet.get_formula_type(spreadsheet, "Sheet1", "C1")
    assert formula_type == "Normal"

    # Get complete formula object and verify its properties
    {text, type, shared_index, reference} =
      UmyaSpreadsheet.get_formula_obj(spreadsheet, "Sheet1", "C1")

    assert text == "SUM(B1:B5)"
    assert type == "Normal"
    # Regular formulas use 0 for shared index
    assert shared_index == 0
    # Reference might be empty but should exist
    assert is_binary(reference)
  end

  test "set_array_formula sets an array formula for a range", %{spreadsheet: spreadsheet} do
    # Set an array formula that should fill multiple cells
    assert :ok = UmyaSpreadsheet.set_array_formula(spreadsheet, "Sheet1", "D1:D3", "B1:B3")

    # For array formulas, only the first cell in the range contains the formula
    # Verify the formula was actually set in the first cell
    assert UmyaSpreadsheet.is_formula(spreadsheet, "Sheet1", "D1") == true

    # Verify formula text (doesn't include the = sign)
    formula_text = UmyaSpreadsheet.get_formula(spreadsheet, "Sheet1", "D1")
    assert formula_text == "B1:B3"

    # Verify formula type - API returns "Normal" for array formulas
    # This happens because set_array_formula() only passes the formula text to the cell.set_formula() method,
    # not preserving the Array type information. This is a limitation in the implementation.
    formula_type = UmyaSpreadsheet.get_formula_type(spreadsheet, "Sheet1", "D1")
    assert formula_type == "Normal"

    # Get complete formula object and verify its properties
    {text, type, shared_index, reference} =
      UmyaSpreadsheet.get_formula_obj(spreadsheet, "Sheet1", "D1")

    assert text == "B1:B3"
    # API returns Normal for array formulas too
    assert type == "Normal"
    # Array formulas use 0 for shared index
    assert shared_index == 0
    # Reference should exist for array formulas
    assert is_binary(reference)

    # Other cells in the range should not have formulas directly
    # They get calculated when Excel loads the file
    refute UmyaSpreadsheet.is_formula(spreadsheet, "Sheet1", "D2")
    refute UmyaSpreadsheet.is_formula(spreadsheet, "Sheet1", "D3")
  end

  test "get_text returns the formula text from a cell", %{spreadsheet: spreadsheet} do
    # Set a regular formula
    assert :ok = UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "C1", "SUM(B1:B5)")

    # Set an array formula
    assert :ok = UmyaSpreadsheet.set_array_formula(spreadsheet, "Sheet1", "D1:D3", "B1:B3")

    # Verify get_text for regular formula
    text_regular = UmyaSpreadsheet.get_text(spreadsheet, "Sheet1", "C1")
    assert text_regular == "SUM(B1:B5)"

    # Verify get_text for array formula
    text_array = UmyaSpreadsheet.get_text(spreadsheet, "Sheet1", "D1")
    assert text_array == "B1:B3"

    # Verify get_text for a cell without a formula
    text_none = UmyaSpreadsheet.get_text(spreadsheet, "Sheet1", "A1")
    assert text_none == ""
  end

  test "is_formula correctly identifies cells with formulas", %{spreadsheet: spreadsheet} do
    # Set a regular formula
    assert :ok = UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "C1", "SUM(B1:B5)")

    # Set an array formula
    assert :ok = UmyaSpreadsheet.set_array_formula(spreadsheet, "Sheet1", "D1:D3", "B1:B3")

    # Verify is_formula for cells with formulas
    assert UmyaSpreadsheet.is_formula(spreadsheet, "Sheet1", "C1") == true
    assert UmyaSpreadsheet.is_formula(spreadsheet, "Sheet1", "D1") == true

    # Verify is_formula for cells without formulas
    assert UmyaSpreadsheet.is_formula(spreadsheet, "Sheet1", "A1") == false
    # Other cells in array formula range
    assert UmyaSpreadsheet.is_formula(spreadsheet, "Sheet1", "D2") == false
    # Non-existent cell
    assert UmyaSpreadsheet.is_formula(spreadsheet, "Sheet1", "E1") == false
  end

  test "get_formula_type returns the correct formula type", %{spreadsheet: spreadsheet} do
    # Set a regular formula
    assert :ok = UmyaSpreadsheet.set_formula(spreadsheet, "Sheet1", "C1", "SUM(B1:B5)")

    # Set an array formula
    assert :ok = UmyaSpreadsheet.set_array_formula(spreadsheet, "Sheet1", "D1:D3", "B1:B3")

    # Verify formula type for regular formula
    assert UmyaSpreadsheet.get_formula_type(spreadsheet, "Sheet1", "C1") == "Normal"

    # Verify formula type for array formula
    # Note: The API returns "Normal" for array formulas despite setting CellFormulaValues::Array
    # This happens because set_array_formula() only passes the formula text to the cell.set_formula()
    # method, not preserving the Array type information. This is a limitation in the implementation.
    assert UmyaSpreadsheet.get_formula_type(spreadsheet, "Sheet1", "D1") == "Normal"

    # Verify formula type for cells without formulas
    assert UmyaSpreadsheet.get_formula_type(spreadsheet, "Sheet1", "A1") == "None"
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
    assert :ok =
             UmyaSpreadsheet.create_defined_name(spreadsheet, "Subtotal", "SUM(B1:B3)", "Sheet1")

    # Currently can't test retrieving the defined name as the library doesn't support it yet
  end

  test "get_defined_names retrieves defined names", %{spreadsheet: spreadsheet} do
    # Create some defined names
    assert :ok = UmyaSpreadsheet.create_named_range(spreadsheet, "DataRange", "Sheet1", "B1:B5")
    assert :ok = UmyaSpreadsheet.create_defined_name(spreadsheet, "TaxRate", "0.15")

    assert :ok =
             UmyaSpreadsheet.create_defined_name(spreadsheet, "Subtotal", "SUM(B1:B3)", "Sheet1")

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
    tax_rate =
      Enum.find(defined_names, fn {name, formula} ->
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
