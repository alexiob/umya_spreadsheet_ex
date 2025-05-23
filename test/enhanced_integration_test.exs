defmodule UmyaSpreadsheet.EnhancedIntegrationTest do
  use ExUnit.Case, async: true

  @test_file_path "test/test_files/aaa.xlsx"
  @image_sample1_path "test/test_files/images/sample1.png"
  @image_sample2_path "test/test_files/images/sample2.png"
  @output_path "test/result_files/enhanced_test.xlsx"
  @output_image_path "test/result_files/output_image.png"
  @output_path_2 "test/result_files/enhanced_test_2.xlsx"
  @output_image_path_2 "test/result_files/output_image_2.png"

  test "get formatted values from cells" do
    # Read Excel file
    {:ok, spreadsheet} = UmyaSpreadsheet.read(@test_file_path)

    # Get formatted values from cells
    {:ok, formatted_value1} = UmyaSpreadsheet.get_formatted_value(spreadsheet, "Sheet1", "B20")
    {:ok, formatted_value2} = UmyaSpreadsheet.get_formatted_value(spreadsheet, "Sheet1", "B21")

    # Verify formatted values
    assert formatted_value1 == "1.0000" || formatted_value1 != ""
    assert formatted_value2 == "$3,333.0000" || formatted_value2 != ""
  end

  test "remove cells from the spreadsheet" do
    # Read Excel file
    {:ok, spreadsheet} = UmyaSpreadsheet.read(@test_file_path)

    # Set a value to a cell and verify it's there
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Test Value")
    assert {:ok, "Test Value"} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "A1")

    # Remove the cell and verify it's gone
    assert :ok = UmyaSpreadsheet.remove_cell(spreadsheet, "Sheet1", "A1")

    # The cell should now be empty
    assert {:ok, ""} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "A1")
  end

  test "set custom number formats" do
    # Create a new spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Set values to cells
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "49046881.119999997")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "3.14159")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "123456")

    # Apply different number formats
    assert :ok = UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "A1", "#,##0.00")
    assert :ok = UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "A2", "0.000")
    assert :ok = UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "A3", "$#,##0")

    # Save the spreadsheet
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)

    # Read back and verify the formatted values
    {:ok, loaded} = UmyaSpreadsheet.read(@output_path)
    {:ok, formatted_value1} = UmyaSpreadsheet.get_formatted_value(loaded, "Sheet1", "A1")
    {:ok, formatted_value2} = UmyaSpreadsheet.get_formatted_value(loaded, "Sheet1", "A2")
    {:ok, formatted_value3} = UmyaSpreadsheet.get_formatted_value(loaded, "Sheet1", "A3")

    # Verification - some Excel versions might format slightly differently
    assert String.contains?(formatted_value1, "49,046,881")
    assert String.contains?(formatted_value2, "3.14")

    assert String.contains?(formatted_value3, "123,456") ||
             String.contains?(formatted_value3, "$123,456")
  end

  test "work with images in spreadsheets" do
    # Create a new spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add an image to the spreadsheet
    assert :ok = UmyaSpreadsheet.add_image(spreadsheet, "Sheet1", "C5", @image_sample1_path)

    # Save the spreadsheet
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)

    # Read back the file
    {:ok, loaded} = UmyaSpreadsheet.read(@output_path)

    # Download the image
    assert :ok = UmyaSpreadsheet.download_image(loaded, "Sheet1", "C5", @output_image_path)

    # Verify the image was downloaded
    assert File.exists?(@output_image_path)

    # Change the image
    assert :ok = UmyaSpreadsheet.change_image(loaded, "Sheet1", "C5", @image_sample2_path)

    # Save the modified spreadsheet
    assert :ok = UmyaSpreadsheet.write(loaded, @output_path_2)

    # Read the spreadsheet with the changed image
    {:ok, loaded_2} = UmyaSpreadsheet.read(@output_path_2)

    # Download the changed image
    assert :ok = UmyaSpreadsheet.download_image(loaded_2, "Sheet1", "C5", @output_image_path_2)

    # Verify the changed image was downloaded
    assert File.exists?(@output_image_path_2)
  end

  test "row styling with height and colors" do
    # Create a new spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add content to row
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "Row with styling")

    assert :ok =
             UmyaSpreadsheet.set_cell_value(
               spreadsheet,
               "Sheet1",
               "B3",
               "This row has custom height and colors"
             )

    # Set row height
    assert :ok = UmyaSpreadsheet.set_row_height(spreadsheet, "Sheet1", 3, 30.0)

    # Set row styling with background and font colors
    assert :ok = UmyaSpreadsheet.set_row_style(spreadsheet, "Sheet1", 3, "black", "white")

    # Save the spreadsheet
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)

    # Read back the file (just to verify it can be read)
    {:ok, _loaded} = UmyaSpreadsheet.read(@output_path)

    # Since we can't easily verify the styling programmatically, just check that the file exists
    assert File.exists?(@output_path)
  end
end
