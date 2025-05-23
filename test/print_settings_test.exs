defmodule UmyaSpreadsheetTest.PrintSettingsTest do
  use ExUnit.Case, async: true

  alias UmyaSpreadsheet

  setup do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()
    %{spreadsheet: spreadsheet}
  end

  test "print settings can be configured", %{spreadsheet: spreadsheet} do
    # Set page orientation to landscape
    assert :ok = UmyaSpreadsheet.set_page_orientation(spreadsheet, "Sheet1", "landscape")

    # Set paper size to A4
    assert :ok = UmyaSpreadsheet.set_paper_size(spreadsheet, "Sheet1", 9)

    # Set page scale to 80%
    assert :ok = UmyaSpreadsheet.set_page_scale(spreadsheet, "Sheet1", 80)

    # Set fit to 1 page wide by 2 pages tall
    assert :ok = UmyaSpreadsheet.set_fit_to_page(spreadsheet, "Sheet1", 1, 2)

    # Set page margins
    assert :ok = UmyaSpreadsheet.set_page_margins(spreadsheet, "Sheet1", 1.0, 0.75, 1.0, 0.75)

    # Set header/footer margins
    assert :ok = UmyaSpreadsheet.set_header_footer_margins(spreadsheet, "Sheet1", 0.5, 0.5)

    # Set header
    assert :ok = UmyaSpreadsheet.set_header(spreadsheet, "Sheet1", "&C&\"Arial,Bold\"Confidential")

    # Set footer
    assert :ok = UmyaSpreadsheet.set_footer(spreadsheet, "Sheet1", "&RPage &P of &N")

    # Center content on page
    assert :ok = UmyaSpreadsheet.set_print_centered(spreadsheet, "Sheet1", true, false)

    # Set print area
    assert :ok = UmyaSpreadsheet.set_print_area(spreadsheet, "Sheet1", "A1:H20")

    # Set print titles (repeating rows/columns)
    assert :ok = UmyaSpreadsheet.set_print_titles(spreadsheet, "Sheet1", "1:2", "A:B")
  end

  # Temporary test to ensure we can run the project
  test "library loads correctly" do
    assert true
  end
end
