defmodule UmyaSpreadsheet.ExtendedProtectionTest do
  use ExUnit.Case, async: true

  @output_path "test/result_files/extended_protection_test.xlsx"
  @temp_output_path "test/result_files/extended_protection_temp.xlsx"
  @password "secret123"

  setup_all do
    # Make sure the result files directory exists
    File.mkdir_p!("test/result_files")
    :ok
  end

  test "create spreadsheet with styled content and password protection" do
    # Create new spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add content with styling
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Confidential Data")

    assert :ok =
             UmyaSpreadsheet.set_cell_value(
               spreadsheet,
               "Sheet1",
               "A2",
               "Protected with password"
             )

    assert :ok = UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "A1", "red")
    assert :ok = UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1", true)
    assert :ok = UmyaSpreadsheet.set_font_color(spreadsheet, "Sheet1", "A2", "blue")

    # Save with password
    assert :ok = UmyaSpreadsheet.write_with_password(spreadsheet, @output_path, @password)

    # Verify file was created
    assert File.exists?(@output_path)

    # File size should be reasonable for an Excel file
    file_size = File.stat!(@output_path).size
    assert file_size > 1000, "File size is too small for an Excel file"
  end

  test "try to read protected file without password should fail" do
    # First make sure we have a protected file to test
    # Create new spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add content with styling
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Test Data")
    assert :ok = UmyaSpreadsheet.write_with_password(spreadsheet, @output_path, @password)

    # Verify file exists
    assert File.exists?(@output_path)

    # Attempt to read without password
    result = UmyaSpreadsheet.read(@output_path)
    # We expect this to fail and return an error
    assert match?({:error, _}, result),
           "Should not be able to read password-protected file without password"
  end

  test "create and verify multiple sheets with protection" do
    # Create new spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add content to first sheet
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Sheet 1 Data")
    assert :ok = UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1", true)

    # Add second sheet
    assert :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "Confidential")

    assert :ok =
             UmyaSpreadsheet.set_cell_value(
               spreadsheet,
               "Confidential",
               "A1",
               "Secret Information"
             )

    assert :ok = UmyaSpreadsheet.set_background_color(spreadsheet, "Confidential", "A1", "red")

    # Save with password
    assert :ok = UmyaSpreadsheet.write_with_password(spreadsheet, @temp_output_path, @password)

    # Verify file was created
    assert File.exists?(@temp_output_path)
  end
end
