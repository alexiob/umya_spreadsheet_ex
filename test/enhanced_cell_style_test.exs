defmodule UmyaSpreadsheet.EnhancedCellStyleTest do
  use ExUnit.Case, async: true

  @output_path "test/result_files/enhanced_cell_style_test.xlsx"
  @password_output_path "test/result_files/enhanced_cell_style_test_password.xlsx"

  setup do
    # Create a new spreadsheet for testing
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Clean up any previous test result files
    File.rm(@output_path)
    File.rm(@password_output_path)

    %{spreadsheet: spreadsheet}
  end

  test "apply and manipulate cell styles", %{spreadsheet: spreadsheet} do
    # Set some cell values to style
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Bold Text")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "Colored Text")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "Different Font")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A4", "Big Text")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A5", "Background Color")

    assert :ok =
             UmyaSpreadsheet.set_cell_value(
               spreadsheet,
               "Sheet1",
               "A6",
               "Wrapped Text Example With Long Text That Should Wrap"
             )

    # Apply styles
    assert :ok = UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1", true)
    assert :ok = UmyaSpreadsheet.set_font_color(spreadsheet, "Sheet1", "A2", "red")
    assert :ok = UmyaSpreadsheet.set_font_name(spreadsheet, "Sheet1", "A3", "Arial")
    assert :ok = UmyaSpreadsheet.set_font_size(spreadsheet, "Sheet1", "A4", 16)
    assert :ok = UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "A5", "blue")
    assert :ok = UmyaSpreadsheet.set_wrap_text(spreadsheet, "Sheet1", "A6", true)
    assert :ok = UmyaSpreadsheet.set_column_width(spreadsheet, "Sheet1", "A", 15.0)

    # Test number formatting
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "1000")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", "0.5")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B3", "12345678")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B4", "0.123456")

    assert :ok = UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "B1", "#,##0.00")
    assert :ok = UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "B2", "0.00%")
    assert :ok = UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "B3", "$#,##0.00")
    assert :ok = UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "B4", "0.000")

    # Test row styling
    assert :ok = UmyaSpreadsheet.set_row_height(spreadsheet, "Sheet1", 7, 30.0)
    assert :ok = UmyaSpreadsheet.set_row_style(spreadsheet, "Sheet1", 7, "black", "white")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A7", "Row with styling")

    assert :ok =
             UmyaSpreadsheet.set_cell_value(
               spreadsheet,
               "Sheet1",
               "B7",
               "White text on black background"
             )

    # Add a new sheet for more tests
    assert :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "MergedCells")

    # Test merged cells
    assert :ok =
             UmyaSpreadsheet.set_cell_value(
               spreadsheet,
               "MergedCells",
               "B2",
               "Merged Cell Content"
             )

    assert :ok = UmyaSpreadsheet.add_merge_cells(spreadsheet, "MergedCells", "B2:D4")
    assert :ok = UmyaSpreadsheet.set_background_color(spreadsheet, "MergedCells", "B2", "yellow")

    # Test column auto-width
    assert :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "AutoWidth")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "AutoWidth", "C1", "Short")

    assert :ok =
             UmyaSpreadsheet.set_cell_value(
               spreadsheet,
               "AutoWidth",
               "C2",
               "This is a longer text to trigger auto-width"
             )

    assert :ok =
             UmyaSpreadsheet.set_cell_value(
               spreadsheet,
               "AutoWidth",
               "C3",
               "Even longer text to demonstrate that auto-width will adjust the column width"
             )

    assert :ok = UmyaSpreadsheet.set_column_auto_width(spreadsheet, "AutoWidth", "C", true)

    # Test sheet visibility
    assert :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "HiddenSheet")

    assert :ok =
             UmyaSpreadsheet.set_cell_value(
               spreadsheet,
               "HiddenSheet",
               "A1",
               "This sheet is hidden"
             )

    assert :ok = UmyaSpreadsheet.set_sheet_state(spreadsheet, "HiddenSheet", "hidden")

    assert :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "VeryHiddenSheet")

    assert :ok =
             UmyaSpreadsheet.set_cell_value(
               spreadsheet,
               "VeryHiddenSheet",
               "A1",
               "This sheet is very hidden"
             )

    assert :ok = UmyaSpreadsheet.set_sheet_state(spreadsheet, "VeryHiddenSheet", "very_hidden")

    # Test grid lines
    assert :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "NoGridLines")

    assert :ok =
             UmyaSpreadsheet.set_cell_value(
               spreadsheet,
               "NoGridLines",
               "A1",
               "This sheet has no grid lines"
             )

    assert :ok = UmyaSpreadsheet.set_show_grid_lines(spreadsheet, "NoGridLines", false)

    # Copy styling between rows and columns
    assert :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "CopyStyles")

    # Setup source row with styling
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "CopyStyles", "A1", "Source Row")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "CopyStyles", "B1", "More content")
    assert :ok = UmyaSpreadsheet.set_font_bold(spreadsheet, "CopyStyles", "A1", true)
    assert :ok = UmyaSpreadsheet.set_background_color(spreadsheet, "CopyStyles", "B1", "green")

    # Target row for copying
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "CopyStyles", "A3", "Target Row")

    assert :ok =
             UmyaSpreadsheet.set_cell_value(spreadsheet, "CopyStyles", "B3", "Will get styling")

    # Copy row styling
    assert :ok = UmyaSpreadsheet.copy_row_styling(spreadsheet, "CopyStyles", 1, 3)

    # Setup source column with styling
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "CopyStyles", "D1", "Source")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "CopyStyles", "D2", "Column")
    assert :ok = UmyaSpreadsheet.set_font_color(spreadsheet, "CopyStyles", "D1", "red")
    assert :ok = UmyaSpreadsheet.set_font_size(spreadsheet, "CopyStyles", "D2", 14)

    # Target column for copying
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "CopyStyles", "F1", "Target")
    assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "CopyStyles", "F2", "Column")

    # Convert column letters to indices
    source_column = column_letter_to_index("D")
    target_column = column_letter_to_index("F")

    # Copy column styling
    assert :ok =
             UmyaSpreadsheet.copy_column_styling(
               spreadsheet,
               "CopyStyles",
               source_column,
               target_column
             )

    # Save the spreadsheet
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)

    # Add password protection and save another copy
    assert :ok = UmyaSpreadsheet.set_workbook_protection(spreadsheet, "test_password")

    assert :ok =
             UmyaSpreadsheet.write_with_password(
               spreadsheet,
               @password_output_path,
               "file_password"
             )

    # Verify files exist
    assert File.exists?(@output_path)
    assert File.exists?(@password_output_path)
  end

  # Helper function to convert column letters to index
  defp column_letter_to_index(column) do
    column
    |> String.to_charlist()
    |> Enum.reduce(0, fn c, acc -> acc * 26 + (c - ?A + 1) end)
  end
end
