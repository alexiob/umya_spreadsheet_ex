defmodule AdvancedCellFormattingTest do
  use ExUnit.Case, async: true

  @underline_output_path "test/result_files/underline_test.xlsx"
  @border_output_path "test/result_files/border_test.xlsx"
  @rotation_output_path "test/result_files/rotation_test.xlsx"
  @indent_output_path "test/result_files/indent_test.xlsx"
  @italic_output_path "test/result_files/italic_test.xlsx"
  @strikethrough_output_path "test/result_files/strikethrough_test.xlsx"

  describe "advanced cell formatting" do
    setup do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add some test data
      :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Italic Text")
      :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "Underlined Text")
      :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "Strikethrough Text")
      :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A4", "Bordered Cell")
      :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A5", "Rotated Text")
      :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A6", "Indented Text")

      %{spreadsheet: spreadsheet}
    end

    test "can set italic formatting", %{spreadsheet: spreadsheet} do
      # Set italic formatting
      assert :ok = UmyaSpreadsheet.set_font_italic(spreadsheet, "Sheet1", "A1", true)

      # We can't directly test the visual format here but we can at least
      # ensure the operation doesn't fail

      # Try to save the file to verify it works
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @italic_output_path)
    end

    test "can set different underline styles", %{spreadsheet: spreadsheet} do
      # Set different underline styles
      assert :ok = UmyaSpreadsheet.set_font_underline(spreadsheet, "Sheet1", "A2", "single")
      assert :ok = UmyaSpreadsheet.set_font_underline(spreadsheet, "Sheet1", "B2", "double")
      assert :ok = UmyaSpreadsheet.set_font_underline(spreadsheet, "Sheet1", "C2", "single_accounting")
      assert :ok = UmyaSpreadsheet.set_font_underline(spreadsheet, "Sheet1", "D2", "double_accounting")
      assert :ok = UmyaSpreadsheet.set_font_underline(spreadsheet, "Sheet1", "E2", "none")

      # Try to save the file to verify it works
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @underline_output_path)
          end

    test "can set strikethrough formatting", %{spreadsheet: spreadsheet} do
      # Set strikethrough formatting
      assert :ok = UmyaSpreadsheet.set_font_strikethrough(spreadsheet, "Sheet1", "A3", true)

      # Try to save the file to verify it works
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @strikethrough_output_path)
    end

    test "can set different border styles", %{spreadsheet: spreadsheet} do
      # Set different border styles
      assert :ok = UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "A4", "all", "thin")
      assert :ok = UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "B4", "left", "medium")
      assert :ok = UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "C4", "top", "thick")
      assert :ok = UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "D4", "right", "dashed")
      assert :ok = UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "E4", "bottom", "dotted")
      assert :ok = UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "F4", "diagonal", "double")

      # Try to save the file to verify it works
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @border_output_path)
    end

    test "can set cell rotation", %{spreadsheet: spreadsheet} do
      # Set cell rotation
      assert :ok = UmyaSpreadsheet.set_cell_rotation(spreadsheet, "Sheet1", "A5", 45)
      assert :ok = UmyaSpreadsheet.set_cell_rotation(spreadsheet, "Sheet1", "B5", -45)
      assert :ok = UmyaSpreadsheet.set_cell_rotation(spreadsheet, "Sheet1", "C5", 90)

      # Try to save the file to verify it works
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @rotation_output_path)
    end

    test "can set cell indentation", %{spreadsheet: spreadsheet} do
      # Set cell indentation
      assert :ok = UmyaSpreadsheet.set_cell_indent(spreadsheet, "Sheet1", "A6", 1)
      assert :ok = UmyaSpreadsheet.set_cell_indent(spreadsheet, "Sheet1", "B6", 2)
      assert :ok = UmyaSpreadsheet.set_cell_indent(spreadsheet, "Sheet1", "C6", 5)

      # Try to save the file to verify it works
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @indent_output_path)
    end

    test "handles errors gracefully", %{spreadsheet: spreadsheet} do
      # Invalid sheet name
      assert {:error, "Sheet not found"} = UmyaSpreadsheet.set_font_italic(spreadsheet, "NonExistentSheet", "A1", true)

      # We don't need to test every function with invalid inputs,
      # since they all use the same error handling mechanism
    end
  end
end
