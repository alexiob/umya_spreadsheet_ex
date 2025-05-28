defmodule UmyaSpreadsheetTest.VmlSupportTest do
  use ExUnit.Case, async: false

  alias UmyaSpreadsheet

  test "create and manipulate VML shapes" do
    # Create a new spreadsheet
    {:ok, book} = UmyaSpreadsheet.new()

    # Add a worksheet
    :ok = UmyaSpreadsheet.add_sheet(book, "TestSheet")

    # Test creating a VML shape
    assert :ok = UmyaSpreadsheet.VmlDrawing.create_shape(book, "TestSheet", "1")

    # Test setting VML shape properties
    assert :ok = UmyaSpreadsheet.VmlDrawing.set_shape_type(book, "TestSheet", "1", "oval")
    assert :ok = UmyaSpreadsheet.VmlDrawing.set_shape_filled(book, "TestSheet", "1", true)

    assert :ok =
             UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(book, "TestSheet", "1", "#FF0000")

    assert :ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroked(book, "TestSheet", "1", true)

    assert :ok =
             UmyaSpreadsheet.VmlDrawing.set_shape_stroke_color(book, "TestSheet", "1", "#000000")

    assert :ok = UmyaSpreadsheet.VmlDrawing.set_shape_stroke_weight(book, "TestSheet", "1", "2pt")

    assert :ok =
             UmyaSpreadsheet.VmlDrawing.set_shape_style(
               book,
               "TestSheet",
               "1",
               "position:absolute;left:100pt;top:50pt;width:200pt;height:100pt"
             )
  end

  test "VML shape validation" do
    # Create a new spreadsheet
    {:ok, book} = UmyaSpreadsheet.new()
    :ok = UmyaSpreadsheet.add_sheet(book, "TestSheet")

    # Test validation errors
    assert {:error, _} = UmyaSpreadsheet.VmlDrawing.create_shape(book, "TestSheet", "")
    assert {:error, _} = UmyaSpreadsheet.VmlDrawing.create_shape(book, "NonExistentSheet", "1")

    # Create a valid shape first
    :ok = UmyaSpreadsheet.VmlDrawing.create_shape(book, "TestSheet", "1")

    # Test property validation
    assert {:error, _} =
             UmyaSpreadsheet.VmlDrawing.set_shape_type(book, "TestSheet", "1", "invalid_type")

    assert {:error, _} =
             UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(
               book,
               "TestSheet",
               "1",
               "invalid_color"
             )

    assert {:error, _} =
             UmyaSpreadsheet.VmlDrawing.set_shape_stroke_color(
               book,
               "TestSheet",
               "1",
               "invalid_color"
             )

    assert {:error, _} =
             UmyaSpreadsheet.VmlDrawing.set_shape_stroke_weight(book, "TestSheet", "1", "")

    assert {:error, _} = UmyaSpreadsheet.VmlDrawing.set_shape_style(book, "TestSheet", "1", "")

    # Test with non-existent shape ID
    assert {:error, _} =
             UmyaSpreadsheet.VmlDrawing.set_shape_type(book, "TestSheet", "999", "rect")
  end

  test "VML shape type validation" do
    {:ok, book} = UmyaSpreadsheet.new()
    :ok = UmyaSpreadsheet.add_sheet(book, "TestSheet")
    :ok = UmyaSpreadsheet.VmlDrawing.create_shape(book, "TestSheet", "1")

    # Test valid shape types
    valid_types = ["rect", "oval", "line", "polyline", "roundrect", "arc"]

    for shape_type <- valid_types do
      assert :ok = UmyaSpreadsheet.VmlDrawing.set_shape_type(book, "TestSheet", "1", shape_type)
    end
  end

  test "VML color format validation" do
    {:ok, book} = UmyaSpreadsheet.new()
    :ok = UmyaSpreadsheet.add_sheet(book, "TestSheet")
    :ok = UmyaSpreadsheet.VmlDrawing.create_shape(book, "TestSheet", "1")

    # Test valid color formats
    assert :ok =
             UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(book, "TestSheet", "1", "#FF0000")

    assert :ok =
             UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(
               book,
               "TestSheet",
               "1",
               "rgb(255,0,0)"
             )

    assert :ok =
             UmyaSpreadsheet.VmlDrawing.set_shape_stroke_color(book, "TestSheet", "1", "#00FF00")

    assert :ok =
             UmyaSpreadsheet.VmlDrawing.set_shape_stroke_color(
               book,
               "TestSheet",
               "1",
               "rgb(0,255,0)"
             )

    # Test invalid color formats
    assert {:error, _} =
             UmyaSpreadsheet.VmlDrawing.set_shape_fill_color(book, "TestSheet", "1", "red")

    assert {:error, _} =
             UmyaSpreadsheet.VmlDrawing.set_shape_stroke_color(book, "TestSheet", "1", "blue")
  end
end
