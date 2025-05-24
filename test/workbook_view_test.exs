defmodule UmyaSpreadsheetTest.WorkbookViewTest do
  use ExUnit.Case, async: true
  doctest UmyaSpreadsheet

  @temp_file "test/result_files/workbook_view_test.xlsx"

  setup do
    # Create a new spreadsheet with multiple sheets for each test
    {:ok, spreadsheet} = UmyaSpreadsheet.new()
    :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "Sheet2")
    :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "Sheet3")

    # Add some data to each sheet for visual reference
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "This is Sheet 1")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet2", "A1", "This is Sheet 2")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet3", "A1", "This is Sheet 3")

    {:ok, %{spreadsheet: spreadsheet}}
  end

  describe "workbook view functions" do
    test "set_active_tab sets which tab is active when opening the workbook", %{spreadsheet: spreadsheet} do
      # Test setting the second tab (Sheet2) as the active tab
      assert :ok = UmyaSpreadsheet.set_active_tab(spreadsheet, 1)

      # Write file to verify changes
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @temp_file)

      # Read back the file and verify that it opens correctly
      assert {:ok, _} = UmyaSpreadsheet.read(@temp_file)
    end

    test "set_workbook_window_position sets window position and size", %{spreadsheet: spreadsheet} do
      # Test setting the window position and size
      assert :ok = UmyaSpreadsheet.set_workbook_window_position(spreadsheet, 100, 50, 800, 600)

      # Write file to verify changes
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @temp_file)

      # Read back the file and verify that it opens correctly
      assert {:ok, _} = UmyaSpreadsheet.read(@temp_file)
    end

    test "combining workbook view settings", %{spreadsheet: spreadsheet} do
      # Apply multiple settings - set Sheet3 active and position the window
      assert :ok = UmyaSpreadsheet.set_active_tab(spreadsheet, 2)
      assert :ok = UmyaSpreadsheet.set_workbook_window_position(spreadsheet, 200, 100, 1024, 768)

      # Also apply some sheet-specific settings
      assert :ok = UmyaSpreadsheet.set_tab_color(spreadsheet, "Sheet1", "#FF0000") # Red tab
      assert :ok = UmyaSpreadsheet.set_tab_color(spreadsheet, "Sheet2", "#00FF00") # Green tab
      assert :ok = UmyaSpreadsheet.set_tab_color(spreadsheet, "Sheet3", "#0000FF") # Blue tab

      # Write file to verify changes
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @temp_file)

      # Read back the file and verify that it opens correctly
      assert {:ok, _} = UmyaSpreadsheet.read(@temp_file)
    end
  end
end
