defmodule UmyaSpreadsheet.ProtectionTest do
  use ExUnit.Case, async: true

  @output_path "test/result_files/protected_test.xlsx"
  @password "secretpassword123"

  setup_all do
    # Make sure the result files directory exists
    File.mkdir_p!("test/result_files")
    :ok
  end

  test "create and save password-protected spreadsheet" do
    # Create new spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add content
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Confidential Data")

    assert :ok =
             UmyaSpreadsheet.set_cell_value(
               spreadsheet,
               "Sheet1",
               "A2",
               "Protected with password"
             )

    assert :ok = UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "A1", "red")

    # Save with password
    assert :ok = UmyaSpreadsheet.write_with_password(spreadsheet, @output_path, @password)

    # Verify file was created
    assert File.exists?(@output_path)

    # File size should be reasonable for an Excel file
    file_size = File.stat!(@output_path).size
    assert file_size > 1000, "File size is too small for an Excel file"
  end
end
