defmodule UmyaSpreadsheet.CsvExportTest do
  use ExUnit.Case, async: true

  @test_file_path "test/test_files/aaa.xlsx"
  @output_csv_path "test/result_files/exported_sheet.csv"
  @output_light_path "test/result_files/exported_light.xlsx"
  @output_pwd_light_path "test/result_files/secure_light.xlsx"

  setup_all do
    # Make sure the result files directory exists
    File.mkdir_p!("test/result_files")
    :ok
  end

  test "export spreadsheet sheet to CSV" do
    # Read Excel file
    {:ok, spreadsheet} = UmyaSpreadsheet.read(@test_file_path)

    # Export Sheet1 to CSV
    result = UmyaSpreadsheet.write_csv(spreadsheet, "Sheet1", @output_csv_path)
    assert result == :ok

    # Verify the CSV file was created
    assert File.exists?(@output_csv_path)

    # Read the CSV content to verify it has data
    csv_content = File.read!(@output_csv_path)
    assert String.length(csv_content) > 0
    # Basic CSV format check
    assert String.contains?(csv_content, ",")
  end

  test "create new empty spreadsheet and add sheets" do
    # Create an empty spreadsheet without any sheets
    {:ok, spreadsheet} = UmyaSpreadsheet.new_empty()

    # Add a new sheet
    :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "CustomSheet")

    # Verify sheet was added
    sheet_names = UmyaSpreadsheet.get_sheet_names(spreadsheet)
    assert "CustomSheet" in sheet_names

    # Add some data
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "CustomSheet", "A1", "Test Data")

    # Save the file
    result = UmyaSpreadsheet.write(spreadsheet, @output_light_path)
    assert result == :ok

    # Verify the file was created
    assert File.exists?(@output_light_path)
  end

  test "using light writer functions" do
    # Create a new spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add some data
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Light Writer Test")

    # Save with light writer
    result = UmyaSpreadsheet.write_light(spreadsheet, @output_light_path)
    assert result == :ok
    assert File.exists?(@output_light_path)

    # Save with password and light writer
    result =
      UmyaSpreadsheet.write_with_password_light(spreadsheet, @output_pwd_light_path, "test123")

    assert result == :ok
    assert File.exists?(@output_pwd_light_path)
  end
end
