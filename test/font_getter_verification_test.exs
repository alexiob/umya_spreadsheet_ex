defmodule FontGetterVerificationTest do
  use ExUnit.Case, async: true

  @output_path "test/result_files/font_getter_verification.xlsx"

  setup_all do
    # Make sure the result files directory exists
    File.mkdir_p!("test/result_files")
    :ok
  end

  describe "Font Getter Functions" do
    test "font getter functions return correct values after setting properties and saving/reading file" do
      # Create a new spreadsheet
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Set various font properties
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Test Font Properties")

      # Set font properties
      UmyaSpreadsheet.set_font_name(spreadsheet, "Sheet1", "A1", "Arial")
      UmyaSpreadsheet.set_font_size(spreadsheet, "Sheet1", "A1", 14)
      UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1", true)
      UmyaSpreadsheet.set_font_italic(spreadsheet, "Sheet1", "A1", true)
      UmyaSpreadsheet.set_font_underline(spreadsheet, "Sheet1", "A1", "single")
      UmyaSpreadsheet.set_font_strikethrough(spreadsheet, "Sheet1", "A1", true)
      UmyaSpreadsheet.set_font_family(spreadsheet, "Sheet1", "A1", "swiss")
      UmyaSpreadsheet.set_font_scheme(spreadsheet, "Sheet1", "A1", "major")
      UmyaSpreadsheet.set_font_color(spreadsheet, "Sheet1", "A1", "FF0000")

      # Add a cell with default formatting for comparison
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Default Cell")

      # Save the file
      UmyaSpreadsheet.write(spreadsheet, @output_path)
      assert File.exists?(@output_path)

      # Read the file back
      {:ok, loaded_spreadsheet} = UmyaSpreadsheet.read(@output_path)

      # Verify cell values were preserved
      {:ok, cell_value_a1} = UmyaSpreadsheet.get_cell_value(loaded_spreadsheet, "Sheet1", "A1")
      assert cell_value_a1 == "Test Font Properties"

      {:ok, cell_value_b1} = UmyaSpreadsheet.get_cell_value(loaded_spreadsheet, "Sheet1", "B1")
      assert cell_value_b1 == "Default Cell"

      # Test getter functions on the loaded spreadsheet
      {:ok, font_name} = UmyaSpreadsheet.get_font_name(loaded_spreadsheet, "Sheet1", "A1")
      assert font_name == "Arial"

      {:ok, font_size} = UmyaSpreadsheet.get_font_size(loaded_spreadsheet, "Sheet1", "A1")
      assert font_size == 14.0

      {:ok, font_bold} = UmyaSpreadsheet.get_font_bold(loaded_spreadsheet, "Sheet1", "A1")
      assert font_bold == true

      {:ok, font_italic} = UmyaSpreadsheet.get_font_italic(loaded_spreadsheet, "Sheet1", "A1")
      assert font_italic == true

      {:ok, font_underline} =
        UmyaSpreadsheet.get_font_underline(loaded_spreadsheet, "Sheet1", "A1")

      assert font_underline == "single"

      {:ok, font_strikethrough} =
        UmyaSpreadsheet.get_font_strikethrough(loaded_spreadsheet, "Sheet1", "A1")

      assert font_strikethrough == true

      {:ok, font_family} = UmyaSpreadsheet.get_font_family(loaded_spreadsheet, "Sheet1", "A1")
      assert font_family == "swiss"

      {:ok, font_scheme} = UmyaSpreadsheet.get_font_scheme(loaded_spreadsheet, "Sheet1", "A1")
      assert font_scheme == "major"

      {:ok, font_color} = UmyaSpreadsheet.get_font_color(loaded_spreadsheet, "Sheet1", "A1")
      # Color should contain the red color, might include alpha channel
      assert String.contains?(font_color, "FF0000") or font_color == "FFFF0000"

      # Test with a cell that has default formatting (should return defaults)
      {:ok, default_font_name} = UmyaSpreadsheet.get_font_name(loaded_spreadsheet, "Sheet1", "B1")
      assert default_font_name == "Calibri"

      {:ok, default_font_size} = UmyaSpreadsheet.get_font_size(loaded_spreadsheet, "Sheet1", "B1")
      assert default_font_size == 11.0

      {:ok, default_font_bold} = UmyaSpreadsheet.get_font_bold(loaded_spreadsheet, "Sheet1", "B1")
      assert default_font_bold == false

      {:ok, default_font_italic} =
        UmyaSpreadsheet.get_font_italic(loaded_spreadsheet, "Sheet1", "B1")

      assert default_font_italic == false

      {:ok, default_font_underline} =
        UmyaSpreadsheet.get_font_underline(loaded_spreadsheet, "Sheet1", "B1")

      assert default_font_underline == "none"

      {:ok, default_font_strikethrough} =
        UmyaSpreadsheet.get_font_strikethrough(loaded_spreadsheet, "Sheet1", "B1")

      assert default_font_strikethrough == false

      {:ok, default_font_family} =
        UmyaSpreadsheet.get_font_family(loaded_spreadsheet, "Sheet1", "B1")

      assert default_font_family == "auto"

      {:ok, default_font_scheme} =
        UmyaSpreadsheet.get_font_scheme(loaded_spreadsheet, "Sheet1", "B1")

      # Default scheme might be "none" or "minor" depending on implementation
      assert default_font_scheme == "minor" or default_font_scheme == "none"

      # Test with a non-existing cell (should return defaults)
      {:ok, empty_font_name} = UmyaSpreadsheet.get_font_name(loaded_spreadsheet, "Sheet1", "Z99")
      assert empty_font_name == "Calibri"

      {:ok, empty_font_size} = UmyaSpreadsheet.get_font_size(loaded_spreadsheet, "Sheet1", "Z99")
      assert empty_font_size == 11.0

      {:ok, empty_font_bold} = UmyaSpreadsheet.get_font_bold(loaded_spreadsheet, "Sheet1", "Z99")
      assert empty_font_bold == false
    end
  end
end
