defmodule UmyaSpreadsheet.StylingTest do
  use ExUnit.Case, async: true

  @output_path "test/result_files/styled_test.xlsx"

  setup_all do
    # Make sure the result files directory exists
    File.mkdir_p!("test/result_files")
    :ok
  end

  test "font styling functions work" do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Test set_font_bold
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Bold Text")
    assert :ok = UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1", true)

    # Test set_font_size
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "Large Text")
    assert :ok = UmyaSpreadsheet.set_font_size(spreadsheet, "Sheet1", "A2", 16)

    # Test set_font_color
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "Colored Text")
    assert :ok = UmyaSpreadsheet.set_font_color(spreadsheet, "Sheet1", "A3", "blue")
  end

  test "create and save styled spreadsheet" do
    # Create new spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Write values to cells
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Hello")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "World")

    # Apply styling
    assert :ok = UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "A1", "red")
    assert :ok = UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1", true)
    assert :ok = UmyaSpreadsheet.set_font_size(spreadsheet, "Sheet1", "B1", 14)
    assert :ok = UmyaSpreadsheet.set_font_color(spreadsheet, "Sheet1", "B1", "blue")

    # Add a new sheet
    assert :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "Sheet2")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet2", "A1", "This is Sheet2")

    # Save the spreadsheet
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)

    # Verify file was created
    assert File.exists?(@output_path)
  end

  test "read back styled spreadsheet" do
    # Create a file first to ensure it exists
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add some content
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Hello")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "World")
    assert :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "Sheet2")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet2", "A1", "This is Sheet2")

    # Save it
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)

    # Verify file exists
    assert File.exists?(@output_path)

    # Read the file we created
    {:ok, loaded} = UmyaSpreadsheet.read(@output_path)

    # Verify cell values
    assert {:ok, "Hello"} = UmyaSpreadsheet.get_cell_value(loaded, "Sheet1", "A1")
    assert {:ok, "World"} = UmyaSpreadsheet.get_cell_value(loaded, "Sheet1", "B1")
    assert {:ok, "This is Sheet2"} = UmyaSpreadsheet.get_cell_value(loaded, "Sheet2", "A1")

    # Verify sheet names
    sheet_names = UmyaSpreadsheet.get_sheet_names(loaded)
    assert Enum.member?(sheet_names, "Sheet1")
    assert Enum.member?(sheet_names, "Sheet2")
    assert length(sheet_names) == 2
  end
end
