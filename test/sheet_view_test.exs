defmodule UmyaSpreadsheetTest.SheetViewTest do
  use ExUnit.Case, async: true
  doctest UmyaSpreadsheet

  @temp_file "test/result_files/sheet_view_test.xlsx"

  setup do
    # Create a new spreadsheet for each test
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    {:ok, %{spreadsheet: spreadsheet}}
  end

  describe "sheet view functions" do
    test "set_tab_color sets the color of sheet tabs", %{spreadsheet: spreadsheet} do
      # Test setting tab color
      assert :ok = UmyaSpreadsheet.set_tab_color(spreadsheet, "Sheet1", "#FF0000")

      # Write file to verify changes
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @temp_file)

      # Read back the file and verify that it opens correctly
      assert {:ok, _} = UmyaSpreadsheet.read(@temp_file)
    end

    test "freeze_panes freezes rows and columns", %{spreadsheet: spreadsheet} do
      # Test freezing the first row
      assert :ok = UmyaSpreadsheet.freeze_panes(spreadsheet, "Sheet1", 1, 0)

      # Write file to verify changes
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @temp_file)

      # Read back the file and verify that it opens correctly
      assert {:ok, _} = UmyaSpreadsheet.read(@temp_file)
    end

    test "split_panes splits the view", %{spreadsheet: spreadsheet} do
      # Test splitting the view
      assert :ok = UmyaSpreadsheet.split_panes(spreadsheet, "Sheet1", 2000, 2000)

      # Write file to verify changes
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @temp_file)

      # Read back the file and verify that it opens correctly
      assert {:ok, _} = UmyaSpreadsheet.read(@temp_file)
    end

    test "set_sheet_view changes the view type", %{spreadsheet: spreadsheet} do
      # Test setting different view types
      assert :ok = UmyaSpreadsheet.set_sheet_view(spreadsheet, "Sheet1", "normal")
      assert :ok = UmyaSpreadsheet.set_sheet_view(spreadsheet, "Sheet1", "page_layout")
      assert :ok = UmyaSpreadsheet.set_sheet_view(spreadsheet, "Sheet1", "page_break_preview")

      # Write file to verify changes
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @temp_file)

      # Read back the file and verify that it opens correctly
      assert {:ok, _} = UmyaSpreadsheet.read(@temp_file)
    end

    test "set_zoom_scale sets the zoom factor", %{spreadsheet: spreadsheet} do
      # Test setting zoom scale
      assert :ok = UmyaSpreadsheet.set_zoom_scale(spreadsheet, "Sheet1", 150)

      # Write file to verify changes
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @temp_file)

      # Read back the file and verify that it opens correctly
      assert {:ok, _} = UmyaSpreadsheet.read(@temp_file)
    end

    test "set_zoom_scale_normal sets normal view zoom factor", %{spreadsheet: spreadsheet} do
      # Test setting normal view zoom scale
      assert :ok = UmyaSpreadsheet.set_zoom_scale_normal(spreadsheet, "Sheet1", 75)

      # Write file to verify changes
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @temp_file)

      # Read back the file and verify that it opens correctly
      assert {:ok, _} = UmyaSpreadsheet.read(@temp_file)
    end

    test "set_zoom_scale_page_layout sets page layout zoom factor", %{spreadsheet: spreadsheet} do
      # Test setting page layout view zoom scale
      assert :ok = UmyaSpreadsheet.set_zoom_scale_page_layout(spreadsheet, "Sheet1", 120)

      # Write file to verify changes
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @temp_file)

      # Read back the file and verify that it opens correctly
      assert {:ok, _} = UmyaSpreadsheet.read(@temp_file)
    end

    test "set_zoom_scale_page_break sets page break preview zoom factor", %{spreadsheet: spreadsheet} do
      # Test setting page break preview zoom scale
      assert :ok = UmyaSpreadsheet.set_zoom_scale_page_break(spreadsheet, "Sheet1", 80)

      # Write file to verify changes
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @temp_file)

      # Read back the file and verify that it opens correctly
      assert {:ok, _} = UmyaSpreadsheet.read(@temp_file)
    end

    test "combined sheet view settings", %{spreadsheet: spreadsheet} do
      # Add some data
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Title")
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "Row 1")
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "Row 2")
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Column 1")
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C1", "Column 2")

      # Apply multiple sheet view settings
      assert :ok = UmyaSpreadsheet.set_tab_color(spreadsheet, "Sheet1", "#00FF00")
      assert :ok = UmyaSpreadsheet.freeze_panes(spreadsheet, "Sheet1", 1, 0)
      assert :ok = UmyaSpreadsheet.set_sheet_view(spreadsheet, "Sheet1", "page_layout")
      assert :ok = UmyaSpreadsheet.set_zoom_scale_page_layout(spreadsheet, "Sheet1", 150)

      # Write file to verify changes
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @temp_file)

      # Read back the file and verify that it opens correctly
      assert {:ok, _} = UmyaSpreadsheet.read(@temp_file)
    end

    test "set_selection sets the active cell and selection range", %{spreadsheet: spreadsheet} do
      # Add some data for visual reference
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Top Left")
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C3", "Selected Cell")
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D5", "Bottom Right")

      # Test setting active cell and selection range
      assert :ok = UmyaSpreadsheet.set_selection(spreadsheet, "Sheet1", "C3", "C3:D5")

      # Write file to verify changes
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @temp_file)

      # Read back the file and verify that it opens correctly
      assert {:ok, _} = UmyaSpreadsheet.read(@temp_file)
    end
  end
end
