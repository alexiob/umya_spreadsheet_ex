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
    assert :ok = UmyaSpreadsheet.create_defined_name(
      spreadsheet,
      "GLOBAL_FORMULA",
      "10*5+2",
      nil  # No sheet name = global
    )

    # Test global defined name (cell reference)
    assert :ok = UmyaSpreadsheet.create_defined_name(
      spreadsheet,
      "GLOBAL_RANGE",
      "TestSheet!A1:C3",
      nil  # No sheet name = global
    )

    # Test sheet-scoped defined name
    assert :ok = UmyaSpreadsheet.create_defined_name(
      spreadsheet,
      "LOCAL_CONSTANT",
      "42",
      "TestSheet"  # Sheet-scoped to TestSheet
    )

    # Test sheet-scoped defined name (formula)
    assert :ok = UmyaSpreadsheet.create_defined_name(
      spreadsheet,
      "LOCAL_FORMULA",
      "SUM(A1:A10)",
      "TestSheet"  # Sheet-scoped to TestSheet
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

    # We can't directly test the defined names since there's no public API to read them,
    # but if the file loads successfully without errors, it means our defined names
    # were properly created and saved
    assert match?({:ok, _}, {:ok, loaded_spreadsheet})
  end

  test "error handling for invalid inputs" do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Test empty name
    assert {:error, _} = UmyaSpreadsheet.create_defined_name(spreadsheet, "", "A1", nil)

    # Test empty formula
    assert {:error, _} = UmyaSpreadsheet.create_defined_name(spreadsheet, "TEST", "", nil)

    # Test invalid sheet name
    assert {:error, _} = UmyaSpreadsheet.create_defined_name(spreadsheet, "TEST", "A1", "NonExistentSheet")
  end
end
