defmodule UmyaSpreadsheetTest.DataValidationTest do
  use ExUnit.Case, async: true

  alias UmyaSpreadsheet

  setup do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()
    %{spreadsheet: spreadsheet}
  end

  test "add_list_validation adds dropdown validation", %{spreadsheet: spreadsheet} do
    # Add list validation to a range
    assert :ok =
             UmyaSpreadsheet.add_list_validation(
               spreadsheet,
               "Sheet1",
               "A1:A5",
               ["Option 1", "Option 2", "Option 3"],
               true,
               "Invalid Selection",
               "Please select a valid option from the list",
               "Selection Required",
               "Please select an option"
             )

    # Add a value to the first cell
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Option 1")
  end

  test "add_number_validation with between operator", %{spreadsheet: spreadsheet} do
    # Add number validation to a range
    assert :ok =
             UmyaSpreadsheet.add_number_validation(
               spreadsheet,
               "Sheet1",
               "B1:B5",
               "between",
               1.0,
               10.0,
               true,
               "Invalid Number",
               "Please enter a number between 1 and 10",
               "Number Input",
               "Enter a number between 1 and 10"
             )

    # Add a valid value
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "5")
  end

  test "add_date_validation with greater than operator", %{spreadsheet: spreadsheet} do
    today = Date.utc_today() |> Date.to_iso8601()

    # Add date validation to a range
    assert :ok =
             UmyaSpreadsheet.add_date_validation(
               spreadsheet,
               "Sheet1",
               "C1:C5",
               "greaterThan",
               today,
               nil,
               true,
               "Invalid Date",
               "Please enter a future date",
               "Date Input",
               "Enter a date in the future"
             )

    # Tomorrow's date (this is just for testing, actual validation happens in Excel)
    tomorrow =
      Date.utc_today()
      |> Date.add(1)
      |> Date.to_iso8601()

    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C1", tomorrow)
  end

  test "add_text_length_validation with maximum length", %{spreadsheet: spreadsheet} do
    # Add text length validation to a range
    assert :ok =
             UmyaSpreadsheet.add_text_length_validation(
               spreadsheet,
               "Sheet1",
               "D1:D5",
               "lessThanOrEqual",
               10,
               nil,
               true,
               "Text Too Long",
               "Please enter at most 10 characters",
               "Text Input",
               "Enter up to 10 characters"
             )

    # Add a valid value
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D1", "Hello")
  end

  test "add_custom_validation with formula", %{spreadsheet: spreadsheet} do
    # Add custom validation to a range
    assert :ok =
             UmyaSpreadsheet.add_custom_validation(
               spreadsheet,
               "Sheet1",
               "E1:E5",
               "MOD(E1,2)=0",
               true,
               "Invalid Value",
               "Please enter an even number",
               "Even Number Required",
               "Enter an even number"
             )

    # Add a valid value
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "E1", "4")
  end

  test "remove_data_validation removes validation rules", %{spreadsheet: spreadsheet} do
    # Add validation first
    assert :ok =
             UmyaSpreadsheet.add_list_validation(
               spreadsheet,
               "Sheet1",
               "F1:F5",
               ["Option 1", "Option 2", "Option 3"]
             )

    # Now remove it
    assert :ok =
             UmyaSpreadsheet.remove_data_validation(
               spreadsheet,
               "Sheet1",
               "F1:F5"
             )

    # We can still set any value now (actual validation would be enforced by Excel)
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "F1", "Other Value")
  end

  test "full validation example with file save/read", %{spreadsheet: spreadsheet} do
    # Add different types of validation
    UmyaSpreadsheet.add_list_validation(
      spreadsheet,
      "Sheet1",
      "A1:A5",
      ["Red", "Green", "Blue"]
    )

    UmyaSpreadsheet.add_number_validation(
      spreadsheet,
      "Sheet1",
      "B1:B5",
      "between",
      1.0,
      100.0
    )

    UmyaSpreadsheet.add_date_validation(
      spreadsheet,
      "Sheet1",
      "C1:C5",
      "greaterThanOrEqual",
      Date.utc_today() |> Date.to_iso8601()
    )

    UmyaSpreadsheet.add_text_length_validation(
      spreadsheet,
      "Sheet1",
      "D1:D5",
      "lessThanOrEqual",
      20
    )

    UmyaSpreadsheet.add_custom_validation(
      spreadsheet,
      "Sheet1",
      "E1:E5",
      "MOD(E1,3)=0"
    )

    # Set some values
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Red")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "50")

    UmyaSpreadsheet.set_cell_value(
      spreadsheet,
      "Sheet1",
      "C1",
      Date.utc_today() |> Date.add(5) |> Date.to_iso8601()
    )

    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D1", "Test text")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "E1", "9")

    # Save the file to verify validations persist
    output_path = "test/result_files/data_validation_full_example.xlsx"
    File.mkdir_p!("test/result_files")
    assert :ok = UmyaSpreadsheet.write(spreadsheet, output_path)
    assert File.exists?(output_path)

    # Read the file back to verify validations and values were saved correctly
    {:ok, loaded_spreadsheet} = UmyaSpreadsheet.read(output_path)

    # Verify the cell values were preserved
    assert {:ok, "Red"} = UmyaSpreadsheet.get_cell_value(loaded_spreadsheet, "Sheet1", "A1")
    assert {:ok, "50"} = UmyaSpreadsheet.get_cell_value(loaded_spreadsheet, "Sheet1", "B1")
    assert {:ok, "Test text"} = UmyaSpreadsheet.get_cell_value(loaded_spreadsheet, "Sheet1", "D1")
    assert {:ok, "9"} = UmyaSpreadsheet.get_cell_value(loaded_spreadsheet, "Sheet1", "E1")

    # Verify that the validations were saved by checking they exist
    {:ok, validations} =
      UmyaSpreadsheet.DataValidation.get_data_validations(loaded_spreadsheet, "Sheet1")

    assert is_list(validations)
    assert length(validations) == 5

    # Verify each validation type is present
    validation_types = Enum.map(validations, & &1.rule_type)
    assert :list in validation_types
    assert :decimal in validation_types or :whole in validation_types
    assert :date in validation_types
    assert :textLength in validation_types
    assert :custom in validation_types

    # Clean up the test file
    File.rm(output_path)
  end
end
