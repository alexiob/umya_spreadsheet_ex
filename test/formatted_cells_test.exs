defmodule UmyaSpreadsheet.FormattedCellsTest do
  use ExUnit.Case, async: true

  @output_path "test/result_files/formatted_cells_test.xlsx"
  setup do
    # Create a new spreadsheet for each test
    {:ok, spreadsheet} = UmyaSpreadsheet.new()
    %{spreadsheet: spreadsheet}
  end

  test "formatted value retrieval", %{spreadsheet: spreadsheet} do
    # Set a number with number format
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "1234.5678")
    UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "A1", "#,##0.00")

    # Get the formatted value
    {:ok, formatted_value} = UmyaSpreadsheet.get_formatted_value(spreadsheet, "Sheet1", "A1")
    assert formatted_value == "1,234.57"

    # Set a date with date format
    # Excel date format
    date_value = Float.to_string(:os.system_time(:second) / 86400 + 25569)
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", date_value)
    UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "A2", "yyyy-mm-dd")

    # The formatted value should be a date string
    {:ok, formatted_date} = UmyaSpreadsheet.get_formatted_value(spreadsheet, "Sheet1", "A2")
    assert String.match?(formatted_date, ~r/\d{4}-\d{2}-\d{2}/)
  end

  test "cell removal", %{spreadsheet: spreadsheet} do
    # Set a value
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Test Value")
    {:ok, value} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "A1")
    assert value == "Test Value"

    # Remove the cell
    UmyaSpreadsheet.remove_cell(spreadsheet, "Sheet1", "A1")

    # Value should be empty after removal
    {:ok, empty_value} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "A1")
    assert empty_value == ""
  end

  test "row styling", %{spreadsheet: spreadsheet} do
    # Set style to a row
    UmyaSpreadsheet.set_row_style(spreadsheet, "Sheet1", 1, "black", "white")

    # Set values in that row
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Styled Cell")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Another Cell")

    # Verify styling was applied to row 1
    {:ok, bg_color_a1} = UmyaSpreadsheet.get_cell_background_color(spreadsheet, "Sheet1", "A1")
    {:ok, font_color_a1} = UmyaSpreadsheet.get_cell_foreground_color(spreadsheet, "Sheet1", "A1")

    # Background should be black, font should be white (or their hex equivalents)
    assert bg_color_a1 == "black" || bg_color_a1 == "FF000000" ||
             String.contains?(bg_color_a1, "000000")

    assert font_color_a1 == "white" || font_color_a1 == "FF000000" ||
             font_color_a1 == "FFFFFF" || String.contains?(font_color_a1, "FFFFFF")

    # Copy styling to another row
    UmyaSpreadsheet.copy_row_styling(spreadsheet, "Sheet1", 1, 2)

    # Verify row 2 received the styling
    {:ok, bg_color_a2} = UmyaSpreadsheet.get_cell_background_color(spreadsheet, "Sheet1", "A2")

    assert bg_color_a2 == "black" || bg_color_a2 == "FF000000" ||
             String.contains?(bg_color_a2, "000000")

    # Partial styling copy
    # Only copy A column style
    UmyaSpreadsheet.copy_row_styling(spreadsheet, "Sheet1", 1, 3, 1, 1)

    # Verify column A in row 3 got the style
    {:ok, bg_color_a3} = UmyaSpreadsheet.get_cell_background_color(spreadsheet, "Sheet1", "A3")

    assert bg_color_a3 == "black" || bg_color_a3 == "FF000000" ||
             String.contains?(bg_color_a3, "000000")

    # Verify column B in row 3 did NOT get the style (should have default background)
    {:ok, bg_color_b3} = UmyaSpreadsheet.get_cell_background_color(spreadsheet, "Sheet1", "B3")

    assert bg_color_b3 == "FFFFFFFF" || bg_color_b3 == "white" ||
             !String.contains?(bg_color_b3, "000000")
  end

  test "column width manipulation", %{spreadsheet: spreadsheet} do
    # Set column width
    UmyaSpreadsheet.set_column_width(spreadsheet, "Sheet1", "A", 20.0)

    # Set auto width
    UmyaSpreadsheet.set_column_auto_width(spreadsheet, "Sheet1", "B", true)
  end

  test "text wrapping", %{spreadsheet: spreadsheet} do
    # Set a multi-line text
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Line 1\nLine 2\nLine 3")

    # Enable text wrapping
    UmyaSpreadsheet.set_wrap_text(spreadsheet, "Sheet1", "A1", true)
  end

  test "sheet and workbook protection", %{spreadsheet: spreadsheet} do
    # Protect sheet without password
    UmyaSpreadsheet.set_sheet_protection(spreadsheet, "Sheet1", nil, true)

    # Protect sheet with password
    UmyaSpreadsheet.set_sheet_protection(spreadsheet, "Sheet1", "password123", true)

    # Protect workbook
    UmyaSpreadsheet.set_workbook_protection(spreadsheet, "secure_password")
  end

  test "sheet visibility and grid lines", %{spreadsheet: spreadsheet} do
    # Add a new sheet
    UmyaSpreadsheet.add_sheet(spreadsheet, "HiddenSheet")

    # Hide the sheet
    UmyaSpreadsheet.set_sheet_state(spreadsheet, "HiddenSheet", "hidden")

    # Hide grid lines
    UmyaSpreadsheet.set_show_grid_lines(spreadsheet, "Sheet1", false)
  end

  test "merge cells", %{spreadsheet: spreadsheet} do
    # Set values
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Merged Cell Content")

    # Merge cells
    UmyaSpreadsheet.add_merge_cells(spreadsheet, "Sheet1", "A1:C3")
  end

  test "creating and saving a comprehensive spreadsheet", %{spreadsheet: spreadsheet} do
    # Set various cell values
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Formatting Test Sheet")
    UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1", true)
    UmyaSpreadsheet.set_font_size(spreadsheet, "Sheet1", "A1", 16)
    UmyaSpreadsheet.add_merge_cells(spreadsheet, "Sheet1", "A1:E1")

    # Number formatting
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "Numbers")
    UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A3", true)

    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B3", "Format")
    UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "B3", true)

    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C3", "Result")
    UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "C3", true)

    # Set number examples
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A4", "1234.56")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B4", "#,##0.00")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A5", "0.42")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B5", "0%")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A6", "12.3")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B6", "$#,##0.00")

    # Apply formats
    UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "C4", "#,##0.00")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C4", "1234.56")

    UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "C5", "0%")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C5", "0.42")

    UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "C6", "$#,##0.00")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C6", "12.3")

    # Styling
    UmyaSpreadsheet.set_row_style(spreadsheet, "Sheet1", 3, "#4472C4", "white")
    UmyaSpreadsheet.set_column_width(spreadsheet, "Sheet1", "A", 15.0)
    UmyaSpreadsheet.set_column_width(spreadsheet, "Sheet1", "B", 15.0)
    UmyaSpreadsheet.set_column_width(spreadsheet, "Sheet1", "C", 15.0)

    # Add a chart
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A10", "Jan")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A11", "Feb")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A12", "Mar")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A13", "Apr")

    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B10", "100")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B11", "150")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B12", "130")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B13", "190")

    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C10", "80")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C11", "120")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C12", "140")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C13", "160")

    data_series = ["Sheet1!$B$10:$B$13", "Sheet1!$C$10:$C$13"]
    series_titles = ["Revenue", "Expenses"]
    point_titles = ["Jan", "Feb", "Mar", "Apr"]

    UmyaSpreadsheet.add_chart(
      spreadsheet,
      "Sheet1",
      "LineChart",
      "E5",
      "K15",
      "Monthly Financial Data",
      data_series,
      series_titles,
      point_titles
    )

    # Add a second sheet and clone first sheet
    UmyaSpreadsheet.add_sheet(spreadsheet, "Sheet2")
    UmyaSpreadsheet.clone_sheet(spreadsheet, "Sheet1", "Sheet1 Copy")

    # Hide a sheet
    UmyaSpreadsheet.set_sheet_state(spreadsheet, "Sheet2", "hidden")

    # Try to save the file - this is just to verify no errors are thrown
    # We don't actually save since this is a test
    UmyaSpreadsheet.write(spreadsheet, @output_path)
  end
end
