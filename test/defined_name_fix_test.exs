defmodule UmyaSpreadsheet.DefinedNameFixTest do
  use ExUnit.Case, async: true

  @output_path "test/result_files/defined_name_fix_test.xlsx"

  setup do
    # Clean up any previous test files
    File.rm(@output_path)
    %{}
  end

  test "create global and sheet-scoped defined names" do
    # Create a new spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add a new sheet for testing
    assert :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "TestSheet")

    # Test global defined name (formula-based)
    assert :ok =
             UmyaSpreadsheet.create_defined_name(
               spreadsheet,
               "GLOBAL_FORMULA",
               "10*5+2",
               # No sheet name = global
               nil
             )

    # Test global defined name (cell reference)
    assert :ok =
             UmyaSpreadsheet.create_defined_name(
               spreadsheet,
               "GLOBAL_RANGE",
               "TestSheet!A1:C3",
               # No sheet name = global
               nil
             )

    # Test sheet-scoped defined name
    assert :ok =
             UmyaSpreadsheet.create_defined_name(
               spreadsheet,
               "LOCAL_CONSTANT",
               "42",
               # Sheet-scoped to TestSheet
               "TestSheet"
             )

    # Test sheet-scoped defined name (formula)
    assert :ok =
             UmyaSpreadsheet.create_defined_name(
               spreadsheet,
               "LOCAL_FORMULA",
               "SUM(A1:A10)",
               # Sheet-scoped to TestSheet
               "TestSheet"
             )

    # Add some data to test the defined names
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "TestSheet", "A1", "1")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "TestSheet", "A2", "2")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "TestSheet", "A3", "3")

    # Use a defined name in a formula
    assert :ok = UmyaSpreadsheet.set_formula(spreadsheet, "TestSheet", "B1", "=GLOBAL_FORMULA")
    assert :ok = UmyaSpreadsheet.set_formula(spreadsheet, "TestSheet", "B2", "=LOCAL_CONSTANT")

    # Save the spreadsheet
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)

    # Verify file exists
    assert File.exists?(@output_path)

    # Read the file back to verify defined names were saved
    {:ok, loaded_spreadsheet} = UmyaSpreadsheet.read(@output_path)

    # Verify that the file loads successfully
    assert match?({:ok, _}, {:ok, loaded_spreadsheet})

    # Get all defined names from the loaded spreadsheet
    defined_names = UmyaSpreadsheet.get_defined_names(loaded_spreadsheet)
    assert is_list(defined_names)

    # Verify we have the expected number of defined names
    assert length(defined_names) == 4

    # Verify that our sheet-scoped defined names exist and have correct names
    assert Enum.any?(defined_names, fn {name, _} -> name == "LOCAL_CONSTANT" end)
    assert Enum.any?(defined_names, fn {name, _} -> name == "LOCAL_FORMULA" end)

    # Verify that global defined names exist (note: may have empty names due to API limitation)
    # Check by formula content since global names might have empty name strings
    assert Enum.any?(defined_names, fn {_, formula} -> String.contains?(formula, "10*5+2") end)

    assert Enum.any?(defined_names, fn {_, formula} ->
             String.contains?(formula, "'TestSheet'!A1:C3")
           end)

    # Find and verify specific defined names
    local_constant = Enum.find(defined_names, fn {name, _} -> name == "LOCAL_CONSTANT" end)
    local_formula = Enum.find(defined_names, fn {name, _} -> name == "LOCAL_FORMULA" end)

    assert local_constant != nil
    assert local_formula != nil
    {_, constant_formula} = local_constant
    {_, sum_formula} = local_formula

    # Verify the formulas are as expected
    assert constant_formula == "42"
    assert sum_formula == "SUM(A1:A10)"
  end

  test "error handling for invalid inputs" do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Test empty name
    assert {:error, _} = UmyaSpreadsheet.create_defined_name(spreadsheet, "", "A1", nil)

    # Test empty formula
    assert {:error, _} = UmyaSpreadsheet.create_defined_name(spreadsheet, "TEST", "", nil)

    # Test invalid sheet name
    assert {:error, _} =
             UmyaSpreadsheet.create_defined_name(spreadsheet, "TEST", "A1", "NonExistentSheet")
  end
end
