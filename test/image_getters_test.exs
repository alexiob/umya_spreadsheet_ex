defmodule UmyaSpreadsheet.ImageGettersTest do
  use ExUnit.Case, async: true

  @output_path "test/result_files/test_image_getters_output.xlsx"
  @image_path_1 "test/test_files/images/sample1.png"
  @image_path_2 "test/test_files/images/sample2.png"

  setup do
    # Create a new spreadsheet for testing
    {:ok, spreadsheet} = UmyaSpreadsheet.new()
    :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "ImageGettersSheet")

    # Add a couple of images to the sheet
    :ok = UmyaSpreadsheet.add_image(spreadsheet, "ImageGettersSheet", "B2", @image_path_1)
    :ok = UmyaSpreadsheet.add_image(spreadsheet, "ImageGettersSheet", "D5", @image_path_2)

    # Clean up any previous test result files
    File.rm(@output_path)

    %{spreadsheet: spreadsheet}
  end

  test "get image dimensions", %{spreadsheet: spreadsheet} do
    # Get dimensions of the first image
    {:ok, {width, height}} =
      UmyaSpreadsheet.get_image_dimensions(spreadsheet, "ImageGettersSheet", "B2")

    # Verify that we got positive dimensions
    assert is_integer(width)
    assert is_integer(height)
    assert width > 0
    assert height > 0

    # The actual size depends on the test image, but we can check it's reasonable
    assert width > 10
    assert width < 2000
    assert height > 10
    assert height < 2000
  end

  test "list all images in a sheet", %{spreadsheet: spreadsheet} do
    # List all images in the sheet
    {:ok, images} = UmyaSpreadsheet.list_images(spreadsheet, "ImageGettersSheet")

    # We should have 2 images
    assert length(images) == 2

    # Check their coordinates
    coordinates = Enum.map(images, fn {coord, _name} -> coord end)
    assert "B2" in coordinates
    assert "D5" in coordinates

    # Check that image names match the expected file names
    image_names = Enum.map(images, fn {_coord, name} -> name end)
    assert Enum.all?(image_names, &String.contains?(&1, "sample"))
  end

  test "get image info", %{spreadsheet: spreadsheet} do
    # Get info about the first image
    {:ok, {name, position, width, height}} =
      UmyaSpreadsheet.get_image_info(spreadsheet, "ImageGettersSheet", "B2")

    # Verify the position
    assert position == "B2"

    # Verify it contains the expected name
    assert String.contains?(name, "sample")

    # Verify dimensions are reasonable
    assert is_integer(width)
    assert is_integer(height)
    assert width > 0
    assert height > 0
  end

  test "error on non-existent sheet", %{spreadsheet: spreadsheet} do
    # Try to get image dimensions from a non-existent sheet
    result = UmyaSpreadsheet.get_image_dimensions(spreadsheet, "NonExistentSheet", "A1")
    assert result == {:error, "Sheet not found"}

    # Try to list images in a non-existent sheet
    result = UmyaSpreadsheet.list_images(spreadsheet, "NonExistentSheet")
    assert result == {:error, "Sheet not found"}

    # Try to get image info from a non-existent sheet
    result = UmyaSpreadsheet.get_image_info(spreadsheet, "NonExistentSheet", "A1")
    assert result == {:error, "Sheet not found"}
  end

  test "error on non-existent image", %{spreadsheet: spreadsheet} do
    # Try to get dimensions of a non-existent image
    result = UmyaSpreadsheet.get_image_dimensions(spreadsheet, "ImageGettersSheet", "Z99")
    assert result == {:error, "Image not found"}

    # Try to get info about a non-existent image
    result = UmyaSpreadsheet.get_image_info(spreadsheet, "ImageGettersSheet", "Z99")
    assert result == {:error, "Image not found"}
  end

  test "write and read back image information", %{spreadsheet: spreadsheet} do
    # First, write the spreadsheet to a file
    :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)

    # Then, read it back
    {:ok, read_spreadsheet} = UmyaSpreadsheet.read(@output_path)

    # List images in the read spreadsheet
    {:ok, images} = UmyaSpreadsheet.list_images(read_spreadsheet, "ImageGettersSheet")

    # We should still have 2 images
    assert length(images) == 2

    # Their coordinates should still be correct
    coordinates = Enum.map(images, fn {coord, _name} -> coord end)
    assert "B2" in coordinates
    assert "D5" in coordinates
  end
end
