defmodule UmyaSpreadsheet.ExtendedStylingTest do
  use ExUnit.Case, async: true

  @output_path "test/result_files/extended_styling_test.xlsx"

  setup_all do
    # Make sure the result files directory exists
    File.mkdir_p!("test/result_files")
    :ok
  end

  test "create spreadsheet with all styling options" do
    # Create new spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Test bold styling
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Bold Text")
    assert :ok = UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1", true)

    # Test font size
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "Large Text")
    assert :ok = UmyaSpreadsheet.set_font_size(spreadsheet, "Sheet1", "A2", 16)

    # Test font color
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "Red Text")
    assert :ok = UmyaSpreadsheet.set_font_color(spreadsheet, "Sheet1", "A3", "red")

    # Test background color
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A4", "Yellow Background")
    assert :ok = UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "A4", "yellow")

    # Test combined styling
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Combined Styling")
    assert :ok = UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "B1", true)
    assert :ok = UmyaSpreadsheet.set_font_size(spreadsheet, "Sheet1", "B1", 14)
    assert :ok = UmyaSpreadsheet.set_font_color(spreadsheet, "Sheet1", "B1", "green")
    assert :ok = UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "B1", "#CCCCCC")

    # Save the spreadsheet
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)

    # Verify file was created
    assert File.exists?(@output_path)

    # Read back and verify
    {:ok, loaded} = UmyaSpreadsheet.read(@output_path)

    # Verify cell values
    assert {:ok, "Bold Text"} = UmyaSpreadsheet.get_cell_value(loaded, "Sheet1", "A1")
    assert {:ok, "Large Text"} = UmyaSpreadsheet.get_cell_value(loaded, "Sheet1", "A2")
    assert {:ok, "Red Text"} = UmyaSpreadsheet.get_cell_value(loaded, "Sheet1", "A3")
    assert {:ok, "Yellow Background"} = UmyaSpreadsheet.get_cell_value(loaded, "Sheet1", "A4")
    assert {:ok, "Combined Styling"} = UmyaSpreadsheet.get_cell_value(loaded, "Sheet1", "B1")
  end
end
