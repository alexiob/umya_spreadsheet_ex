defmodule UmyaSpreadsheetTest.SheetViewGetterTest do
  use ExUnit.Case, async: true
  doctest UmyaSpreadsheet

  @temp_file "test/result_files/sheet_view_getter_test.xlsx"

  setup do
    # Create a new spreadsheet for each test
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    {:ok, %{spreadsheet: spreadsheet}}
  end

  describe "sheet view getter functions" do
    test "get_show_grid_lines returns grid lines setting", %{spreadsheet: spreadsheet} do
      # Default should be true
      assert {:ok, true} = UmyaSpreadsheet.get_show_grid_lines(spreadsheet, "Sheet1")

      # Set to false and verify
      :ok = UmyaSpreadsheet.set_show_grid_lines(spreadsheet, "Sheet1", false)
      assert {:ok, false} = UmyaSpreadsheet.get_show_grid_lines(spreadsheet, "Sheet1")

      # Set back to true and verify
      :ok = UmyaSpreadsheet.set_show_grid_lines(spreadsheet, "Sheet1", true)
      assert {:ok, true} = UmyaSpreadsheet.get_show_grid_lines(spreadsheet, "Sheet1")
    end

    test "get_zoom_scale returns zoom scale", %{spreadsheet: spreadsheet} do
      # Default should be 100
      assert {:ok, 100} = UmyaSpreadsheet.get_zoom_scale(spreadsheet, "Sheet1")

      # Set to 150 and verify
      :ok = UmyaSpreadsheet.set_zoom_scale(spreadsheet, "Sheet1", 150)
      assert {:ok, 150} = UmyaSpreadsheet.get_zoom_scale(spreadsheet, "Sheet1")

      # Set to 75 and verify
      :ok = UmyaSpreadsheet.set_zoom_scale(spreadsheet, "Sheet1", 75)
      assert {:ok, 75} = UmyaSpreadsheet.get_zoom_scale(spreadsheet, "Sheet1")
    end

    test "get_tab_color returns tab color", %{spreadsheet: spreadsheet} do
      # Default should be no color (empty string)
      assert {:ok, ""} = UmyaSpreadsheet.get_tab_color(spreadsheet, "Sheet1")

      # Set to red and verify
      :ok = UmyaSpreadsheet.set_tab_color(spreadsheet, "Sheet1", "#FF0000")
      assert {:ok, "#FF0000"} = UmyaSpreadsheet.get_tab_color(spreadsheet, "Sheet1")

      # Set to blue and verify
      :ok = UmyaSpreadsheet.set_tab_color(spreadsheet, "Sheet1", "#0000FF")
      assert {:ok, "#0000FF"} = UmyaSpreadsheet.get_tab_color(spreadsheet, "Sheet1")
    end

    test "get_sheet_view returns sheet view type", %{spreadsheet: spreadsheet} do
      # Default should be normal view
      assert {:ok, "normal"} = UmyaSpreadsheet.get_sheet_view(spreadsheet, "Sheet1")

      # Set to page layout and verify
      :ok = UmyaSpreadsheet.set_sheet_view(spreadsheet, "Sheet1", "page_layout")
      assert {:ok, "page_layout"} = UmyaSpreadsheet.get_sheet_view(spreadsheet, "Sheet1")

      # Set to page break preview and verify
      :ok = UmyaSpreadsheet.set_sheet_view(spreadsheet, "Sheet1", "page_break_preview")
      assert {:ok, "page_break_preview"} = UmyaSpreadsheet.get_sheet_view(spreadsheet, "Sheet1")

      # Set back to normal and verify
      :ok = UmyaSpreadsheet.set_sheet_view(spreadsheet, "Sheet1", "normal")
      assert {:ok, "normal"} = UmyaSpreadsheet.get_sheet_view(spreadsheet, "Sheet1")
    end

    test "get_selection returns active cell and selection range", %{spreadsheet: spreadsheet} do
      # Default selection should be A1
      assert {:ok, %{"active_cell" => "A1", "sqref" => "A1"}} =
               UmyaSpreadsheet.get_selection(spreadsheet, "Sheet1")

      # Set selection and verify
      :ok = UmyaSpreadsheet.set_selection(spreadsheet, "Sheet1", "C3", "C3:D5")

      assert {:ok, %{"active_cell" => "C3", "sqref" => "C3:D5"}} =
               UmyaSpreadsheet.get_selection(spreadsheet, "Sheet1")

      # Set different selection and verify
      :ok = UmyaSpreadsheet.set_selection(spreadsheet, "Sheet1", "B2", "B2:E10")

      assert {:ok, %{"active_cell" => "B2", "sqref" => "B2:E10"}} =
               UmyaSpreadsheet.get_selection(spreadsheet, "Sheet1")
    end

    test "persistence of view settings in file", %{spreadsheet: spreadsheet} do
      # Configure settings
      :ok = UmyaSpreadsheet.set_show_grid_lines(spreadsheet, "Sheet1", false)
      :ok = UmyaSpreadsheet.set_zoom_scale(spreadsheet, "Sheet1", 150)
      :ok = UmyaSpreadsheet.set_tab_color(spreadsheet, "Sheet1", "#00FF00")
      :ok = UmyaSpreadsheet.set_sheet_view(spreadsheet, "Sheet1", "page_layout")
      :ok = UmyaSpreadsheet.set_selection(spreadsheet, "Sheet1", "D4", "D4:F6")

      # Write file
      :ok = UmyaSpreadsheet.write(spreadsheet, @temp_file)

      # Read back the file
      {:ok, loaded_spreadsheet} = UmyaSpreadsheet.read(@temp_file)

      # Verify all settings were preserved
      assert {:ok, false} = UmyaSpreadsheet.get_show_grid_lines(loaded_spreadsheet, "Sheet1")
      assert {:ok, 150} = UmyaSpreadsheet.get_zoom_scale(loaded_spreadsheet, "Sheet1")
      assert {:ok, "#00FF00"} = UmyaSpreadsheet.get_tab_color(loaded_spreadsheet, "Sheet1")
      assert {:ok, "page_layout"} = UmyaSpreadsheet.get_sheet_view(loaded_spreadsheet, "Sheet1")

      assert {:ok, %{"active_cell" => "D4", "sqref" => "D4:F6"}} =
               UmyaSpreadsheet.get_selection(loaded_spreadsheet, "Sheet1")
    end
  end
end
