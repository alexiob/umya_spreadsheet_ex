defmodule UmyaSpreadsheetTest.VmlIntegrationTest do
  use ExUnit.Case, async: false

  alias UmyaSpreadsheet

  @test_file_path "test/result_files/vml_shapes_test.xlsx"

  setup do
    # Ensure the result directory exists
    File.mkdir_p!("test/result_files")
    :ok
  end

  test "create spreadsheet with VML shapes and save to file" do
    # Create a new spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add some data to make it interesting
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "VML Shapes Demo")

    :ok =
      UmyaSpreadsheet.set_cell_value(
        spreadsheet,
        "Sheet1",
        "A2",
        "This spreadsheet contains VML shapes"
      )

    # Create various VML shapes

    # Shape 1: Red rectangle
    :ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "Sheet1", "shape1")
    :ok = UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, "Sheet1", "shape1", "rect")

    :ok =
      UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(spreadsheet, "Sheet1", "shape1", "#FF0000")

    :ok =
      UmyaSpreadsheet.VmlDrawing.set_shape_style(
        spreadsheet,
        "Sheet1",
        "shape1",
        "position:absolute;left:50pt;top:50pt;width:100pt;height:60pt"
      )

    # Shape 2: Blue oval
    :ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "Sheet1", "shape2")
    :ok = UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, "Sheet1", "shape2", "oval")

    :ok =
      UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(spreadsheet, "Sheet1", "shape2", "#0000FF")

    :ok =
      UmyaSpreadsheet.VmlDrawing.set_shape_stroke_color(
        spreadsheet,
        "Sheet1",
        "shape2",
        "#000000"
      )

    :ok =
      UmyaSpreadsheet.VmlDrawing.set_shape_stroke_weight(spreadsheet, "Sheet1", "shape2", "2pt")

    :ok =
      UmyaSpreadsheet.VmlDrawing.set_shape_style(
        spreadsheet,
        "Sheet1",
        "shape2",
        "position:absolute;left:200pt;top:50pt;width:80pt;height:80pt"
      )

    # Shape 3: Green rounded rectangle (no fill, just stroke)
    :ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "Sheet1", "shape3")
    :ok = UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, "Sheet1", "shape3", "roundrect")
    :ok = UmyaSpreadsheet.VmlDrawing.set_shape_filled(spreadsheet, "Sheet1", "shape3", false)

    :ok =
      UmyaSpreadsheet.VmlDrawing.set_shape_stroke_color(
        spreadsheet,
        "Sheet1",
        "shape3",
        "#00FF00"
      )

    :ok =
      UmyaSpreadsheet.VmlDrawing.set_shape_stroke_weight(spreadsheet, "Sheet1", "shape3", "3pt")

    :ok =
      UmyaSpreadsheet.VmlDrawing.set_shape_style(
        spreadsheet,
        "Sheet1",
        "shape3",
        "position:absolute;left:50pt;top:150pt;width:120pt;height:40pt"
      )

    # Save the spreadsheet
    :ok = UmyaSpreadsheet.write(spreadsheet, @test_file_path)

    # Verify the file was created
    assert File.exists?(@test_file_path)

    # Verify file size is reasonable (should be larger than empty file)
    %{size: file_size} = File.stat!(@test_file_path)
    assert file_size > 5000, "File size should be reasonable for a spreadsheet with VML shapes"
  end

  test "validate VML shape error handling" do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Test creating shape on non-existent sheet
    assert {:error, error_msg} =
             UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "NonExistent", "1")

    assert error_msg =~ "Failed to create VML shape"

    # Test setting properties on non-existent shape
    assert {:error, error_msg} =
             UmyaSpreadsheet.VmlDrawing.set_shape_type(spreadsheet, "Sheet1", "999", "rect")

    assert error_msg =~ "Failed to set VML shape type"
  end
end
