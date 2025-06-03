defmodule UmyaSpreadsheet.DrawingGettersTest do
  use ExUnit.Case
  doctest UmyaSpreadsheet

  alias UmyaSpreadsheet.{Drawing}

  describe "drawing getters" do
    setup do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Test Drawing Getters")
      {:ok, %{spreadsheet: spreadsheet}}
    end

    test "get_shapes returns shapes correctly", %{spreadsheet: spreadsheet} do
      # Add a shape first
      :ok =
        Drawing.add_shape(
          spreadsheet,
          "Sheet1",
          "B2",
          "rectangle",
          100.0,
          50.0,
          "blue",
          "black",
          1.0
        )

      # Test getting shapes
      {:ok, shapes} = Drawing.get_shapes(spreadsheet, "Sheet1")

      # Verify we got at least one shape back
      assert length(shapes) > 0

      # Verify the shape details
      shape = hd(shapes)
      assert shape.type == "rectangle"
      assert shape.cell == "B2"
      assert shape.width == 100.0
      assert shape.height == 50.0
      # blue
      assert shape.fill_color == "0000FF"
      # black
      assert shape.outline_color == "000000"
      assert shape.outline_width == 1.0

      # Test with cell range filter
      {:ok, filtered_shapes} = Drawing.get_shapes(spreadsheet, "Sheet1", "B2:B2")
      assert length(filtered_shapes) == 1

      # Test with non-matching cell range
      {:ok, no_shapes} = Drawing.get_shapes(spreadsheet, "Sheet1", "C3:C3")
      assert no_shapes == []
    end

    test "get_text_boxes returns text boxes correctly", %{spreadsheet: spreadsheet} do
      # Add a text box first
      :ok =
        Drawing.add_text_box(
          spreadsheet,
          "Sheet1",
          "C3",
          "Hello World",
          150.0,
          75.0,
          "white",
          "black",
          "gray",
          1.0
        )

      # Test getting text boxes
      {:ok, text_boxes} = Drawing.get_text_boxes(spreadsheet, "Sheet1")

      # Verify we got at least one text box back
      assert length(text_boxes) > 0

      # Verify the text box details
      text_box = hd(text_boxes)
      assert text_box.cell == "C3"
      assert text_box.text == "Hello World"
      assert text_box.width == 150.0
      assert text_box.height == 75.0
      # white
      assert text_box.fill_color == "FFFFFF"
      # black
      assert text_box.text_color == "000000"
      # gray (hex representation may vary)
      assert text_box.outline_color in ["808080", "CCCCCC"]
      assert text_box.outline_width == 1.0

      # Test with cell range filter
      {:ok, filtered_text_boxes} = Drawing.get_text_boxes(spreadsheet, "Sheet1", "C3:C3")
      assert length(filtered_text_boxes) == 1

      # Test with non-matching cell range
      {:ok, no_text_boxes} = Drawing.get_text_boxes(spreadsheet, "Sheet1", "D4:D4")
      assert no_text_boxes == []
    end

    test "get_connectors returns connectors correctly", %{spreadsheet: spreadsheet} do
      # Add a connector first
      :ok =
        Drawing.add_connector(
          spreadsheet,
          "Sheet1",
          "D4",
          "E5",
          "red",
          1.5
        )

      # Test getting connectors
      {:ok, connectors} = Drawing.get_connectors(spreadsheet, "Sheet1")

      # Verify we got at least one connector back
      assert length(connectors) > 0

      # Verify the connector details
      connector = hd(connectors)
      assert connector.from_cell == "D4"
      assert connector.to_cell == "E5"
      # red
      assert connector.line_color == "FF0000"
      assert connector.line_width == 1.5

      # Test with cell range filter (should match if either cell is in range)
      {:ok, filtered_connectors} = Drawing.get_connectors(spreadsheet, "Sheet1", "D4:D4")
      assert length(filtered_connectors) == 1

      {:ok, filtered_connectors2} = Drawing.get_connectors(spreadsheet, "Sheet1", "E5:E5")
      assert length(filtered_connectors2) == 1

      # Test with non-matching cell range
      {:ok, no_connectors} = Drawing.get_connectors(spreadsheet, "Sheet1", "F6:F6")
      assert no_connectors == []
    end

    test "has_drawing_objects correctly identifies drawings", %{spreadsheet: spreadsheet} do
      # Initially no drawing objects
      {:ok, has_objects} = Drawing.has_drawing_objects(spreadsheet, "Sheet1")
      assert has_objects == false

      # Add a shape
      :ok =
        Drawing.add_shape(
          spreadsheet,
          "Sheet1",
          "B2",
          "rectangle",
          100.0,
          50.0,
          "blue",
          "black",
          1.0
        )

      # Now should have drawing objects
      {:ok, has_objects} = Drawing.has_drawing_objects(spreadsheet, "Sheet1")
      assert has_objects == true

      # Test with specific cell range
      {:ok, has_objects_in_range} = Drawing.has_drawing_objects(spreadsheet, "Sheet1", "B2:B2")
      assert has_objects_in_range == true

      # Test with non-matching cell range
      {:ok, has_objects_outside_range} =
        Drawing.has_drawing_objects(spreadsheet, "Sheet1", "C3:C3")

      assert has_objects_outside_range == false
    end

    test "count_drawing_objects counts correctly", %{spreadsheet: spreadsheet} do
      # Initially no drawing objects
      {:ok, count} = Drawing.count_drawing_objects(spreadsheet, "Sheet1")
      assert count == 0

      # Add a shape
      :ok =
        Drawing.add_shape(
          spreadsheet,
          "Sheet1",
          "B2",
          "rectangle",
          100.0,
          50.0,
          "blue",
          "black",
          1.0
        )

      # Add a text box
      :ok =
        Drawing.add_text_box(
          spreadsheet,
          "Sheet1",
          "C3",
          "Hello World",
          150.0,
          75.0,
          "white",
          "black",
          "gray",
          1.0
        )

      # Add a connector
      :ok =
        Drawing.add_connector(
          spreadsheet,
          "Sheet1",
          "D4",
          "E5",
          "red",
          1.5
        )

      # Now should have 3 drawing objects
      {:ok, count} = Drawing.count_drawing_objects(spreadsheet, "Sheet1")
      assert count == 3

      # Test with specific cell range
      {:ok, count_in_range} = Drawing.count_drawing_objects(spreadsheet, "Sheet1", "B2:B2")
      assert count_in_range == 1

      # Test with range containing multiple objects
      {:ok, count_in_range} = Drawing.count_drawing_objects(spreadsheet, "Sheet1", "B2:C3")
      assert count_in_range == 2
    end

    test "error cases for non-existent sheet", %{spreadsheet: spreadsheet} do
      {:error, :not_found} = Drawing.get_shapes(spreadsheet, "NonExistentSheet")
      {:error, :not_found} = Drawing.get_text_boxes(spreadsheet, "NonExistentSheet")
      {:error, :not_found} = Drawing.get_connectors(spreadsheet, "NonExistentSheet")
      {:error, :not_found} = Drawing.has_drawing_objects(spreadsheet, "NonExistentSheet")
      {:error, :not_found} = Drawing.count_drawing_objects(spreadsheet, "NonExistentSheet")
    end
  end
end
