defmodule UmyaSpreadsheet.ComprehensiveIntegrationTest do
  use ExUnit.Case, async: true

  @input_path "test/test_files/aaa.xlsx"
  @output_path "test/result_files/comprehensive_output.xlsx"
  @password_output_path "test/result_files/comprehensive_output_password.xlsx"
  @lazy_output_path "test/result_files/comprehensive_output_lazy.xlsx"
  @downloaded_image_path "test/result_files/downloaded_image.png"

  setup do
    # Ensure the output directory exists
    File.mkdir_p!("test/result_files")

    # Clean up any previous output files
    File.rm(@output_path)
    File.rm(@password_output_path)
    File.rm(@lazy_output_path)
    File.rm(@downloaded_image_path)

    %{}
  end

  test "read, modify and write an Excel file" do
    # Read the input file
    assert File.exists?(@input_path), "Input file doesn't exist: #{@input_path}"
    {:ok, spreadsheet} = UmyaSpreadsheet.read(@input_path)

    # Modify cell values
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "TEST1")

    # Read back the value to verify it was set
    {:ok, a1_value} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "A1")
    assert a1_value == "TEST1"

    # Remove a cell
    assert :ok = UmyaSpreadsheet.remove_cell(spreadsheet, "Sheet1", "A1")

    # Verify the cell is now empty
    {:ok, a1_value_after_removal} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "A1")
    assert a1_value_after_removal == ""

    # Download an image from the original file if it exists at position M17
    download_result =
      UmyaSpreadsheet.download_image(spreadsheet, "Sheet1", "M17", @downloaded_image_path)

    # Change image if it exists
    if download_result == :ok do
      # Verify the image exists
      assert File.exists?(@downloaded_image_path)

      # Get a path to a sample image to change it with
      source_image = "test/test_files/images/sample1.png"

      # If the source image exists, use it to change the image
      if File.exists?(source_image) do
        assert :ok = UmyaSpreadsheet.change_image(spreadsheet, "Sheet1", "M17", source_image)
      end
    end

    # Check formatting functionality
    # Get formatted value of cell B20
    {:ok, formatted_value} = UmyaSpreadsheet.get_formatted_value(spreadsheet, "Sheet1", "B20")
    assert String.contains?(formatted_value, ".") || formatted_value == ""

    # Apply custom number format to a cell
    assert :ok =
             UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A10", "49046881.119999997")

    assert :ok = UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "A10", "#,##0.00")

    # Set row height
    assert :ok = UmyaSpreadsheet.set_row_height(spreadsheet, "Sheet1", 3, 46.0)

    # Set font name
    assert :ok = UmyaSpreadsheet.set_font_name(spreadsheet, "Sheet1", "A12", "Arial")

    # Apply background and font colors
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A15", "Styled Cell")
    assert :ok = UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "A15", "black")
    assert :ok = UmyaSpreadsheet.set_font_color(spreadsheet, "Sheet1", "A15", "white")

    # Save the spreadsheet
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)

    # Save a password-protected version
    assert :ok =
             UmyaSpreadsheet.write_with_password(spreadsheet, @password_output_path, "password")

    # Verify the files were created
    assert File.exists?(@output_path)
    assert File.exists?(@password_output_path)
  end

  test "lazy read and write an Excel file" do
    # Read the input file with lazy loading
    assert File.exists?(@input_path), "Input file doesn't exist: #{@input_path}"
    {:ok, spreadsheet} = UmyaSpreadsheet.lazy_read(@input_path)

    # Modify cell values
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "LAZY_TEST")

    # Read back the value to verify it was set
    {:ok, a1_value} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "A1")
    assert a1_value == "LAZY_TEST"

    # Save the spreadsheet
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @lazy_output_path)

    # Verify the file was created
    assert File.exists?(@lazy_output_path)
  end

  test "comprehensive sheet operations" do
    # Create a new spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add new sheets
    assert :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "Sheet2")
    assert :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "Sheet3")

    # Clone a sheet
    assert :ok = UmyaSpreadsheet.clone_sheet(spreadsheet, "Sheet1", "Sheet1 Clone")

    # Add a sheet to be deleted
    assert :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "ToBeDeleted")

    # Verify sheets exist
    sheet_names = UmyaSpreadsheet.get_sheet_names(spreadsheet)
    assert Enum.member?(sheet_names, "Sheet1")
    assert Enum.member?(sheet_names, "Sheet2")
    assert Enum.member?(sheet_names, "Sheet3")
    assert Enum.member?(sheet_names, "Sheet1 Clone")
    assert Enum.member?(sheet_names, "ToBeDeleted")

    # Delete a sheet
    assert :ok = UmyaSpreadsheet.remove_sheet(spreadsheet, "ToBeDeleted")

    # Verify the sheet was deleted
    updated_sheet_names = UmyaSpreadsheet.get_sheet_names(spreadsheet)
    refute Enum.member?(updated_sheet_names, "ToBeDeleted")

    # Add data to test row/column operations
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet2", "A1", "Original")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet2", "B1", "Row")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet2", "A2", "Below")

    # Insert new rows
    assert :ok = UmyaSpreadsheet.insert_new_row(spreadsheet, "Sheet2", 2, 2)

    # Now the "Below" cell should be at A4
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet2", "A2", "Inserted Row 1")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet2", "A3", "Inserted Row 2")

    # Insert new columns
    assert :ok = UmyaSpreadsheet.insert_new_column(spreadsheet, "Sheet2", "B", 2)

    # Add content to the inserted columns
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet2", "B1", "Col 1")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet2", "C1", "Col 2")

    # Move a range of cells
    assert :ok = UmyaSpreadsheet.move_range(spreadsheet, "Sheet2", "A1:D4", 2, 2)

    # Set column width
    assert :ok = UmyaSpreadsheet.set_column_width(spreadsheet, "Sheet2", "A", 15.0)
    assert :ok = UmyaSpreadsheet.set_column_auto_width(spreadsheet, "Sheet2", "B", true)

    # Merged cells
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet3", "B2", "Merged Cell")
    assert :ok = UmyaSpreadsheet.add_merge_cells(spreadsheet, "Sheet3", "B2:D4")

    # Sheet protection and visibility
    assert :ok = UmyaSpreadsheet.set_sheet_protection(spreadsheet, "Sheet3", "password", true)
    assert :ok = UmyaSpreadsheet.set_sheet_state(spreadsheet, "Sheet3", "hidden")
    assert :ok = UmyaSpreadsheet.set_show_grid_lines(spreadsheet, "Sheet2", false)

    # Save the spreadsheet
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)

    # Verify the file was created
    assert File.exists?(@output_path)
  end
end
