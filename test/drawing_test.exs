defmodule UmyaSpreadsheet.DrawingTest do
  use ExUnit.Case, async: true

  @path_shapes "test/result_files/test_shapes.xlsx"
  @path_text_box "test/result_files/test_text_box.xlsx"
  @path_connector "test/result_files/test_connector.xlsx"
  @path_flowchart "test/result_files/test_flowchart.xlsx"

  describe "shapes and drawing" do
    test "add_shape adds a shape to the spreadsheet" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      assert :ok =
               UmyaSpreadsheet.add_shape(
                 spreadsheet,
                 "Sheet1",
                 "B3",
                 "rectangle",
                 200,
                 100,
                 "blue",
                 "black",
                 1.0
               )

      # Also test other shape types
      assert :ok =
               UmyaSpreadsheet.add_shape(
                 spreadsheet,
                 "Sheet1",
                 "D3",
                 "ellipse",
                 150,
                 150,
                 "red",
                 "black",
                 1.0
               )

      assert :ok =
               UmyaSpreadsheet.add_shape(
                 spreadsheet,
                 "Sheet1",
                 "F3",
                 "diamond",
                 100,
                 100,
                 "green",
                 "black",
                 1.0
               )

      # Save the file to verify visually
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @path_shapes)
    end

    test "add_text_box adds a text box to the spreadsheet" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      assert :ok =
               UmyaSpreadsheet.add_text_box(
                 spreadsheet,
                 "Sheet1",
                 "B5",
                 "This is a text box\nwith multiple lines",
                 200,
                 100,
                 "white",
                 "black",
                 "#888888",
                 1.0
               )

      # Save the file to verify visually
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @path_text_box)
    end

    test "add_connector adds a connector line between cells" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add some shapes to connect
      assert :ok =
               UmyaSpreadsheet.add_shape(
                 spreadsheet,
                 "Sheet1",
                 "B3",
                 "rectangle",
                 100,
                 50,
                 "blue",
                 "black",
                 1.0
               )

      assert :ok =
               UmyaSpreadsheet.add_shape(
                 spreadsheet,
                 "Sheet1",
                 "E3",
                 "rectangle",
                 100,
                 50,
                 "green",
                 "black",
                 1.0
               )

      # Add connector between the shapes
      assert :ok =
               UmyaSpreadsheet.add_connector(
                 spreadsheet,
                 "Sheet1",
                 "C3",
                 "E3",
                 "red",
                 1.5
               )

      # Save the file to verify visually
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @path_connector)
    end

    test "create a simple flowchart with shapes, text boxes, and connectors" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add flowchart nodes
      assert :ok =
               UmyaSpreadsheet.add_shape(
                 spreadsheet,
                 "Sheet1",
                 "B2",
                 "rectangle",
                 150,
                 80,
                 "#E6F2FF",
                 "black",
                 1.0
               )

      assert :ok =
               UmyaSpreadsheet.add_shape(
                 spreadsheet,
                 "Sheet1",
                 "B6",
                 "diamond",
                 150,
                 100,
                 "#FFE6E6",
                 "black",
                 1.0
               )

      assert :ok =
               UmyaSpreadsheet.add_shape(
                 spreadsheet,
                 "Sheet1",
                 "B10",
                 "rectangle",
                 150,
                 80,
                 "#E6FFE6",
                 "black",
                 1.0
               )

      # Add text for the nodes
      assert :ok =
               UmyaSpreadsheet.add_text_box(
                 spreadsheet,
                 "Sheet1",
                 "B2",
                 "Start",
                 150,
                 80,
                 "transparent",
                 "black",
                 "transparent",
                 0.0
               )

      assert :ok =
               UmyaSpreadsheet.add_text_box(
                 spreadsheet,
                 "Sheet1",
                 "B6",
                 "Decision",
                 150,
                 100,
                 "transparent",
                 "black",
                 "transparent",
                 0.0
               )

      assert :ok =
               UmyaSpreadsheet.add_text_box(
                 spreadsheet,
                 "Sheet1",
                 "B10",
                 "End",
                 150,
                 80,
                 "transparent",
                 "black",
                 "transparent",
                 0.0
               )

      # Add connectors
      assert :ok = UmyaSpreadsheet.add_connector(spreadsheet, "Sheet1", "B4", "B6", "black", 1.0)
      assert :ok = UmyaSpreadsheet.add_connector(spreadsheet, "Sheet1", "B8", "B10", "black", 1.0)

      # Save the file to verify visually
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @path_flowchart)
    end
  end
end
