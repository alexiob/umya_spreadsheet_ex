defmodule FontManagementTest do
  use ExUnit.Case, async: true
  doctest UmyaSpreadsheet

  @output_path "test/result_files/font_management_test.xlsx"

  setup_all do
    # Make sure the result files directory exists
    File.mkdir_p!("test/result_files")
    :ok
  end

  describe "Font Management" do
    test "font properties are correctly applied to Excel file" do
      # Create a new spreadsheet
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # ---- Font Family Tests ----
      # Add test data with different font families
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Default Font")
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "Roman Family (Serif)")
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "Swiss Family (Sans-Serif)")
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A4", "Modern Family (Monospace)")
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A5", "Script Family")
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A6", "Decorative Family")

      # Apply different font families
      # Roman/Serif (1)
      UmyaSpreadsheet.set_font_family(spreadsheet, "Sheet1", "A2", "roman")
      # Swiss/Sans-Serif (2)
      UmyaSpreadsheet.set_font_family(spreadsheet, "Sheet1", "A3", "swiss")
      # Modern/Monospace (3)
      UmyaSpreadsheet.set_font_family(spreadsheet, "Sheet1", "A4", "modern")
      # Script (4)
      UmyaSpreadsheet.set_font_family(spreadsheet, "Sheet1", "A5", "script")
      # Decorative (5)
      UmyaSpreadsheet.set_font_family(spreadsheet, "Sheet1", "A6", "decorative")

      # Set column width for better visibility
      UmyaSpreadsheet.set_column_width(spreadsheet, "Sheet1", "A", 30.0)

      # ---- Font Scheme Tests ----
      # Add test data with different font schemes
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Default Scheme")
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", "Major Font Scheme")
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B3", "Minor Font Scheme")
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B4", "No Font Scheme")

      # Apply different font schemes
      # Major scheme
      UmyaSpreadsheet.set_font_scheme(spreadsheet, "Sheet1", "B2", "major")
      # Minor scheme
      UmyaSpreadsheet.set_font_scheme(spreadsheet, "Sheet1", "B3", "minor")
      # No scheme
      UmyaSpreadsheet.set_font_scheme(spreadsheet, "Sheet1", "B4", "none")

      # Set column width for better visibility
      UmyaSpreadsheet.set_column_width(spreadsheet, "Sheet1", "B", 30.0)

      # ---- Advanced Typography Controls ----
      # Add a title for the test section
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C1", "Advanced Typography Demo")
      UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "C1", true)
      UmyaSpreadsheet.set_font_size(spreadsheet, "Sheet1", "C1", 14)

      # Add test data with typography controls
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C2", "Default Typography")
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C3", "Custom Font with Major Scheme")
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C4", "Sans-serif with Minor Scheme")

      # Apply advanced typography
      # Custom font with Major scheme
      UmyaSpreadsheet.set_font_name(spreadsheet, "Sheet1", "C3", "Cambria")
      UmyaSpreadsheet.set_font_scheme(spreadsheet, "Sheet1", "C3", "major")
      UmyaSpreadsheet.set_font_family(spreadsheet, "Sheet1", "C3", "roman")

      # Sans-serif with Minor scheme
      UmyaSpreadsheet.set_font_name(spreadsheet, "Sheet1", "C4", "Calibri")
      UmyaSpreadsheet.set_font_scheme(spreadsheet, "Sheet1", "C4", "minor")
      UmyaSpreadsheet.set_font_family(spreadsheet, "Sheet1", "C4", "swiss")

      # Set column width for better visibility
      UmyaSpreadsheet.set_column_width(spreadsheet, "Sheet1", "C", 35.0)

      # Add an explicit style property to ensure all cells have styles applied
      # This helps ensure Excel properly includes style information for all cells
      Enum.each(["A1", "A2", "A3", "A4", "A5", "A6"], fn cell ->
        UmyaSpreadsheet.set_font_name(spreadsheet, "Sheet1", cell, "Calibri")
      end)

      # Save the spreadsheet
      UmyaSpreadsheet.write(spreadsheet, @output_path)

      # Verify the file was created
      assert File.exists?(@output_path)
    end
  end
end
