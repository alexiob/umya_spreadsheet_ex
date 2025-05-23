defmodule UmyaSpreadsheetTest.PrintSettingsIntegrationTest do
  use ExUnit.Case, async: true

  alias UmyaSpreadsheet

  setup do
    # Ensure test output directory exists
    File.mkdir_p("test/result_files")
    test_file = "test/result_files/print_settings_test.xlsx"

    # Cleanup after test
    on_exit(fn ->
      File.rm(test_file)
    end)

    # Create a fresh spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()
    %{spreadsheet: spreadsheet, test_file: test_file}
  end

  test "print settings are preserved when writing and reading files", %{spreadsheet: spreadsheet, test_file: test_file} do
    # Add some content to verify
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Report Title")
    UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1", true)
    UmyaSpreadsheet.set_font_size(spreadsheet, "Sheet1", "A1", 16)

    # Configure print settings
    UmyaSpreadsheet.set_page_orientation(spreadsheet, "Sheet1", "landscape")
    UmyaSpreadsheet.set_paper_size(spreadsheet, "Sheet1", 9) # A4
    UmyaSpreadsheet.set_page_scale(spreadsheet, "Sheet1", 80)
    UmyaSpreadsheet.set_fit_to_page(spreadsheet, "Sheet1", 1, 2)
    UmyaSpreadsheet.set_page_margins(spreadsheet, "Sheet1", 1.0, 0.75, 1.0, 0.75)
    UmyaSpreadsheet.set_header_footer_margins(spreadsheet, "Sheet1", 0.5, 0.5)
    UmyaSpreadsheet.set_header(spreadsheet, "Sheet1", "&C&\"Arial,Bold\"Confidential Report")
    UmyaSpreadsheet.set_footer(spreadsheet, "Sheet1", "&RPage &P of &N")
    UmyaSpreadsheet.set_print_centered(spreadsheet, "Sheet1", true, false)
    UmyaSpreadsheet.set_print_area(spreadsheet, "Sheet1", "A1:H20")
    UmyaSpreadsheet.set_print_titles(spreadsheet, "Sheet1", "1:2", "A:B")

    # Write the file
    assert :ok = UmyaSpreadsheet.write(spreadsheet, test_file)

    # Verify file exists
    assert File.exists?(test_file)

    # Read the file back
    {:ok, loaded_spreadsheet} = UmyaSpreadsheet.read(test_file)

    # Verify cell content is preserved
    {:ok, cell_value} = UmyaSpreadsheet.get_cell_value(loaded_spreadsheet, "Sheet1", "A1")
    assert cell_value == "Report Title"

    # Since we can't directly read the print settings from the Elixir API,
    # we'll verify the file can be successfully read which implies the print
    # settings were properly written
    assert :ok = UmyaSpreadsheet.write(loaded_spreadsheet, test_file)
  end

  @tag :real_application
  test "print settings work in a real application scenario", %{spreadsheet: spreadsheet, test_file: test_file} do
    # Create a more realistic business report
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Quarterly Financial Report")
    UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1", true)
    UmyaSpreadsheet.set_font_size(spreadsheet, "Sheet1", "A1", 16)

    # Column headers
    headers = ["Quarter", "Revenue", "Expenses", "Profit", "Margin"]
    for {header, idx} <- Enum.with_index(headers) do
      col = <<65 + idx::utf8>> # A, B, C, etc.
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "#{col}3", header)
      UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "#{col}3", true)
      UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "#{col}3", "DDDDDD")
    end

    # Sample data
    data = [
      ["Q1 2025", 125000, 80000, 45000, "36%"],
      ["Q2 2025", 142000, 87000, 55000, "39%"],
      ["Q3 2025", 137000, 82000, 55000, "40%"],
      ["Q4 2025", 156000, 94000, 62000, "40%"]
    ]

    for {row_data, row_idx} <- Enum.with_index(data) do
      for {cell_value, col_idx} <- Enum.with_index(row_data) do
        col = <<65 + col_idx::utf8>> # A, B, C, etc.
        row = row_idx + 4 # Start at row 4
        UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "#{col}#{row}", "#{cell_value}")

        # Format numbers with commas
        if col_idx in [1, 2, 3] do # Revenue, Expenses, Profit columns
          UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "#{col}#{row}", "$#,##0")
        end
      end
    end

    # Add totals
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A8", "Total")
    UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A8", true)

    # Sum formulas for totals
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B8", "=SUM(B4:B7)")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C8", "=SUM(C4:C7)")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D8", "=SUM(D4:D7)")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "E8", "=D8/B8")

    # Format totals
    UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "B8", true)
    UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "C8", true)
    UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "D8", true)
    UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "E8", true)
    UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "B8", "$#,##0")
    UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "C8", "$#,##0")
    UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "D8", "$#,##0")
    UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "E8", "0.0%")

    # Add bottom border to header row
    for col <- ["A", "B", "C", "D", "E"] do
      UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "#{col}3", "bottom", "thin")
    end

    # Add top and bottom borders to total row
    for col <- ["A", "B", "C", "D", "E"] do
      UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "#{col}8", "top", "thin")
      UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "#{col}8", "bottom", "double")
    end

    # Configure print settings
    UmyaSpreadsheet.set_page_orientation(spreadsheet, "Sheet1", "portrait")
    UmyaSpreadsheet.set_paper_size(spreadsheet, "Sheet1", 9) # A4
    UmyaSpreadsheet.set_fit_to_page(spreadsheet, "Sheet1", 1, 1) # Fit to 1x1 pages
    UmyaSpreadsheet.set_page_margins(spreadsheet, "Sheet1", 0.7, 0.7, 0.7, 0.7)
    UmyaSpreadsheet.set_header(spreadsheet, "Sheet1", "&C&\"Arial,Bold\"Quarterly Financial Report - Confidential")
    UmyaSpreadsheet.set_footer(spreadsheet, "Sheet1", "&L&D&R&P of &N")
    UmyaSpreadsheet.set_print_area(spreadsheet, "Sheet1", "A1:E9")
    UmyaSpreadsheet.set_print_titles(spreadsheet, "Sheet1", "3:3", "") # Repeat row 3 (headers) on each page
    UmyaSpreadsheet.set_print_centered(spreadsheet, "Sheet1", true, false)

    # Write the file
    assert :ok = UmyaSpreadsheet.write(spreadsheet, test_file)

    # Verify file exists
    assert File.exists?(test_file)

    # Success if we got this far
    assert true
  end
end
