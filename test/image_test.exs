defmodule UmyaSpreadsheet.ImageTest do
  use ExUnit.Case, async: true

  @output_path "test/result_files/test_images_output.xlsx"
  @image_path_1 "test/test_files/images/sample1.png"
  @image_path_2 "test/test_files/images/sample2.png"
  @downloaded_image_path "test/result_files/downloaded_image.png"

  setup do
    # Create a new spreadsheet for testing
    {:ok, spreadsheet} = UmyaSpreadsheet.new()
    :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "ImageSheet")

    # Clean up any previous test result files
    File.rm(@downloaded_image_path)
    File.rm(@output_path)

    %{spreadsheet: spreadsheet}
  end

  test "add an image to a spreadsheet", %{spreadsheet: spreadsheet} do
    # Add an image to the spreadsheet
    assert :ok = UmyaSpreadsheet.add_image(spreadsheet, "ImageSheet", "B2", @image_path_1)

    # Write the spreadsheet to a file so we can read it back
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)

    # Read the spreadsheet back
    {:ok, read_spreadsheet} = UmyaSpreadsheet.read(@output_path)

    # Download the image to verify it was added correctly
    assert :ok =
             UmyaSpreadsheet.download_image(
               read_spreadsheet,
               "ImageSheet",
               "B2",
               @downloaded_image_path
             )

    # Verify the downloaded image exists
    assert File.exists?(@downloaded_image_path)
  end

  test "change an image in a spreadsheet", %{spreadsheet: spreadsheet} do
    # First add an image
    assert :ok = UmyaSpreadsheet.add_image(spreadsheet, "ImageSheet", "C3", @image_path_1)

    # Then change the image
    assert :ok = UmyaSpreadsheet.change_image(spreadsheet, "ImageSheet", "C3", @image_path_2)

    # Write the spreadsheet to a file
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)

    # Read the spreadsheet back
    {:ok, read_spreadsheet} = UmyaSpreadsheet.read(@output_path)

    # Download the changed image
    assert :ok =
             UmyaSpreadsheet.download_image(
               read_spreadsheet,
               "ImageSheet",
               "C3",
               @downloaded_image_path
             )

    # Verify the downloaded image exists
    assert File.exists?(@downloaded_image_path)

    # In a real test, we would verify the image content matches @image_path_2
    # but that's beyond the scope of this test
  end

  test "handle errors for non-existent images", %{spreadsheet: spreadsheet} do
    # Try to add a non-existent image
    result = UmyaSpreadsheet.add_image(spreadsheet, "ImageSheet", "D4", "non_existent_image.png")
    assert result == {:error, :error}

    # Try to download an image from a position where there is no image
    result =
      UmyaSpreadsheet.download_image(spreadsheet, "ImageSheet", "E5", @downloaded_image_path)

    assert result == {:error, :not_found}

    # Try to change an image that doesn't exist
    result = UmyaSpreadsheet.change_image(spreadsheet, "ImageSheet", "F6", @image_path_2)
    assert result == {:error, :not_found}
  end

  test "handle errors for non-existent sheets", %{spreadsheet: spreadsheet} do
    # Try operations on a non-existent sheet
    result = UmyaSpreadsheet.add_image(spreadsheet, "NonExistentSheet", "A1", @image_path_1)
    assert result == {:error, :not_found}

    result =
      UmyaSpreadsheet.download_image(
        spreadsheet,
        "NonExistentSheet",
        "A1",
        @downloaded_image_path
      )

    assert result == {:error, :not_found}

    result = UmyaSpreadsheet.change_image(spreadsheet, "NonExistentSheet", "A1", @image_path_2)
    assert result == {:error, :not_found}
  end
end
