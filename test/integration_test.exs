defmodule UmyaSpreadsheet.IntegrationTest do
  use ExUnit.Case, async: true

  @test_file_path "test/test_files/aaa.xlsx"
  @empty_file_path "test/test_files/aaa_empty.xlsx"
  @xlsm_file_path "test/test_files/aaa.xlsm"
  @libre_file_path "test/test_files/libre2.xlsx"

  @result_file_path "test/result_files/result.xlsx"
  @result_password_file_path "test/result_files/result_password.xlsx"
  @result_lazy_file_path "test/result_files/result_lazy.xlsx"
  @result_empty_file_path "test/result_files/result_empty.xlsx"
  @result_xlsm_file_path "test/result_files/result.xlsm"
  @result_lazy_xlsm_file_path "test/result_files/result_lazy.xlsm"
  @result_libre_file_path "test/result_files/result_libre.xlsx"

  @password "secretpassword123"

  setup_all do
    # Make sure the result files directory exists
    File.mkdir_p!("test/result_files")
    :ok
  end

  test "read and write Excel file" do
    # Read Excel file
    {:ok, spreadsheet} = UmyaSpreadsheet.read(@test_file_path)

    # Modify content
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "TEST1")

    # Write back to a new file
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @result_file_path)

    # Verify file was created
    assert File.exists?(@result_file_path)

    # Read back the file and verify content
    {:ok, loaded} = UmyaSpreadsheet.read(@result_file_path)
    assert {:ok, "TEST1"} = UmyaSpreadsheet.get_cell_value(loaded, "Sheet1", "A1")
  end

  test "read and write with password protection" do
    # Read Excel file
    {:ok, spreadsheet} = UmyaSpreadsheet.read(@test_file_path)

    # Modify content
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "PROTECTED")

    # Write back with password protection
    assert :ok =
             UmyaSpreadsheet.write_with_password(
               spreadsheet,
               @result_password_file_path,
               @password
             )

    # Verify file was created
    assert File.exists?(@result_password_file_path)

    # File size should be reasonable
    file_size = File.stat!(@result_password_file_path).size
    assert file_size > 1000, "File size is too small for an Excel file"
  end

  test "lazy read and write" do
    # Read Excel file using lazy loading
    {:ok, spreadsheet} = UmyaSpreadsheet.lazy_read(@test_file_path)

    # Modify content
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "LAZY_TEST")

    # Write back to a new file
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @result_lazy_file_path)

    # Verify file was created
    assert File.exists?(@result_lazy_file_path)

    # Read back the file and verify content
    {:ok, loaded} = UmyaSpreadsheet.read(@result_lazy_file_path)
    assert {:ok, "LAZY_TEST"} = UmyaSpreadsheet.get_cell_value(loaded, "Sheet1", "A1")
  end

  test "read and write empty Excel file" do
    # Read empty Excel file
    {:ok, spreadsheet} = UmyaSpreadsheet.read(@empty_file_path)

    # Write back to a new file
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @result_empty_file_path)

    # Verify file was created
    assert File.exists?(@result_empty_file_path)
  end

  test "read and write XLSM file" do
    # Read XLSM file
    {:ok, spreadsheet} = UmyaSpreadsheet.read(@xlsm_file_path)

    # Modify content
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "XLSM_TEST")

    # Add a new sheet
    assert :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "New Sheet")

    # Write back to a new file
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @result_xlsm_file_path)

    # Verify file was created
    assert File.exists?(@result_xlsm_file_path)

    # Read back the file and verify content
    {:ok, loaded} = UmyaSpreadsheet.read(@result_xlsm_file_path)
    assert {:ok, "XLSM_TEST"} = UmyaSpreadsheet.get_cell_value(loaded, "Sheet1", "A1")

    # Verify sheet names
    sheet_names = UmyaSpreadsheet.get_sheet_names(loaded)
    assert Enum.member?(sheet_names, "Sheet1")
    assert Enum.member?(sheet_names, "New Sheet")
  end

  test "lazy read and write XLSM file" do
    # Read XLSM file using lazy loading
    {:ok, spreadsheet} = UmyaSpreadsheet.lazy_read(@xlsm_file_path)

    # Modify content
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "LAZY_XLSM_TEST")

    # Write back to a new file
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @result_lazy_xlsm_file_path)

    # Verify file was created
    assert File.exists?(@result_lazy_xlsm_file_path)
  end

  test "read and write LibreOffice file" do
    # Read LibreOffice Excel file
    {:ok, spreadsheet} = UmyaSpreadsheet.read(@libre_file_path)

    # Write back to a new file
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @result_libre_file_path)

    # Verify file was created
    assert File.exists?(@result_libre_file_path)
  end

  test "cell styling" do
    {:ok, spreadsheet} = UmyaSpreadsheet.read(@test_file_path)

    # Apply styling to cells
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Styled Text")
    assert :ok = UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "A1", "red")
    assert :ok = UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1", true)
    assert :ok = UmyaSpreadsheet.set_font_size(spreadsheet, "Sheet1", "A2", 16)
    assert :ok = UmyaSpreadsheet.set_font_color(spreadsheet, "Sheet1", "A2", "blue")

    # Write modified file
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @result_file_path)

    # Verify file was created
    assert File.exists?(@result_file_path)
  end

  test "sheet operations" do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add sheets
    assert :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "Sheet2")
    assert :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "Sheet3")

    # Add content to different sheets
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Sheet1 Content")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet2", "A1", "Sheet2 Content")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet3", "A1", "Sheet3 Content")

    # Write to file
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @result_file_path)

    # Read back and verify
    {:ok, loaded} = UmyaSpreadsheet.read(@result_file_path)

    # Verify sheet names
    sheet_names = UmyaSpreadsheet.get_sheet_names(loaded)
    assert Enum.member?(sheet_names, "Sheet1")
    assert Enum.member?(sheet_names, "Sheet2")
    assert Enum.member?(sheet_names, "Sheet3")
    assert length(sheet_names) == 3

    # Verify content in each sheet
    assert {:ok, "Sheet1 Content"} = UmyaSpreadsheet.get_cell_value(loaded, "Sheet1", "A1")
    assert {:ok, "Sheet2 Content"} = UmyaSpreadsheet.get_cell_value(loaded, "Sheet2", "A1")
    assert {:ok, "Sheet3 Content"} = UmyaSpreadsheet.get_cell_value(loaded, "Sheet3", "A1")
  end

  test "move range operation" do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add content
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Move Me")

    # Move content
    assert :ok = UmyaSpreadsheet.move_range(spreadsheet, "Sheet1", "A1:A1", 1, 1)

    # Verify cell values after move
    assert {:ok, ""} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "A1")
    assert {:ok, "Move Me"} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "B2")

    # Write to file
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @result_file_path)

    # Read back and verify
    {:ok, loaded} = UmyaSpreadsheet.read(@result_file_path)
    assert {:ok, "Move Me"} = UmyaSpreadsheet.get_cell_value(loaded, "Sheet1", "B2")
  end
end
