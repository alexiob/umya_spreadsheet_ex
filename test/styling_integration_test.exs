defmodule UmyaSpreadsheet.StylingIntegrationTest do
  use ExUnit.Case, async: true

  @output_path "test/result_files/styling_integration_test.xlsx"

  setup do
    # Create a new spreadsheet for testing
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Clean up any previous test result files
    File.rm(@output_path)

    %{spreadsheet: spreadsheet}
  end

  test "apply various styles to cells", %{spreadsheet: spreadsheet} do
    # Add content to cells
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Cell with font styles")

    :ok =
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "Cell with background color")

    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "Cell with font color")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A4", "Cell with font size")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A5", "Cell with bold font")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A6", "Cell with custom font")

    # Apply font styles
    :ok = UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1", true)
    :ok = UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "A2", "blue")
    :ok = UmyaSpreadsheet.set_font_color(spreadsheet, "Sheet1", "A3", "red")
    :ok = UmyaSpreadsheet.set_font_size(spreadsheet, "Sheet1", "A4", 14)
    :ok = UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A5", true)
    :ok = UmyaSpreadsheet.set_font_name(spreadsheet, "Sheet1", "A6", "Arial")

    # Test copying styles
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Original Row")
    :ok = UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "B1", "yellow")
    :ok = UmyaSpreadsheet.set_font_color(spreadsheet, "Sheet1", "B1", "blue")

    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", "Copied Row Style")
    :ok = UmyaSpreadsheet.copy_row_styling(spreadsheet, "Sheet1", 1, 2)

    # Test column widths and auto-width
    :ok = UmyaSpreadsheet.set_column_width(spreadsheet, "Sheet1", "A", 20.0)
    :ok = UmyaSpreadsheet.set_column_auto_width(spreadsheet, "Sheet1", "B", true)

    # Test text wrapping
    :ok =
      UmyaSpreadsheet.set_cell_value(
        spreadsheet,
        "Sheet1",
        "C1",
        "This text is very long and should wrap to the next line when wrapping is enabled"
      )

    :ok = UmyaSpreadsheet.set_wrap_text(spreadsheet, "Sheet1", "C1", true)
    :ok = UmyaSpreadsheet.set_column_width(spreadsheet, "Sheet1", "C", 15.0)

    # Test merged cells
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D1", "Merged cells region")
    :ok = UmyaSpreadsheet.add_merge_cells(spreadsheet, "Sheet1", "D1:F3")

    # Test sheet visibility and grid lines
    :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "Hidden Sheet")
    :ok = UmyaSpreadsheet.set_sheet_state(spreadsheet, "Hidden Sheet", "hidden")

    :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "No Grid Lines")
    :ok = UmyaSpreadsheet.set_show_grid_lines(spreadsheet, "No Grid Lines", false)

    # Test protection
    :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "Protected Sheet")

    :ok =
      UmyaSpreadsheet.set_sheet_protection(spreadsheet, "Protected Sheet", "password123", true)

    # Save the spreadsheet
    :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)

    # Also test with workbook protection
    :ok = UmyaSpreadsheet.set_workbook_protection(spreadsheet, "workbook_password")

    # Save with a different name
    protected_output_path = "test/result_files/styling_integration_protected_test.xlsx"
    :ok = UmyaSpreadsheet.write(spreadsheet, protected_output_path)

    # Verify the files were created
    assert File.exists?(@output_path)
    assert File.exists?(protected_output_path)

    # Read back the files to verify they can be opened
    {:ok, _read_spreadsheet} = UmyaSpreadsheet.read(@output_path)
  end
end
