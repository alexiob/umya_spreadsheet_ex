defmodule UmyaSpreadsheet.MixedStylingTest do
  use ExUnit.Case, async: true

  @output_path "test/result_files/mixed_styling_test.xlsx"

  setup_all do
    # Make sure the result files directory exists
    File.mkdir_p!("test/result_files")
    :ok
  end

  test "create spreadsheet with mixed styling and verify read back" do
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

    # Read it back
    {:ok, loaded} = UmyaSpreadsheet.read(@output_path)

    # Verify cell values
    assert {:ok, "Hello"} = UmyaSpreadsheet.get_cell_value(loaded, "Sheet1", "A1")
    assert {:ok, "World"} = UmyaSpreadsheet.get_cell_value(loaded, "Sheet1", "B1")
    assert {:ok, "This is Sheet2"} = UmyaSpreadsheet.get_cell_value(loaded, "Sheet2", "A1")

    # Get and verify all sheet names
    sheet_names = UmyaSpreadsheet.get_sheet_names(loaded)
    assert Enum.member?(sheet_names, "Sheet1")
    assert Enum.member?(sheet_names, "Sheet2")
    assert length(sheet_names) == 2
  end
end
