defmodule CellFormattingGettersTest do
  use ExUnit.Case, async: true

  @output_path "test/result_files/cell_formatting_getters.xlsx"

  setup_all do
    # Make sure the result files directory exists
    File.mkdir_p!("test/result_files")
    :ok
  end

  describe "Cell Formatting Getter Functions" do
    test "alignment getter functions return correct values" do
      # Create a new spreadsheet
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Set cell value
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Test Alignment")

      # Set alignment properties using existing setter functions
      UmyaSpreadsheet.set_cell_alignment(spreadsheet, "Sheet1", "A1", "center", "middle")
      # Note: We need to check if these functions exist for wrap_text and text_rotation
      # For now, test with default values

      # Test alignment getters
      {:ok, h_align} = UmyaSpreadsheet.get_cell_horizontal_alignment(spreadsheet, "Sheet1", "A1")
      {:ok, v_align} = UmyaSpreadsheet.get_cell_vertical_alignment(spreadsheet, "Sheet1", "A1")
      {:ok, wrap_text} = UmyaSpreadsheet.get_cell_wrap_text(spreadsheet, "Sheet1", "A1")
      {:ok, rotation} = UmyaSpreadsheet.get_cell_text_rotation(spreadsheet, "Sheet1", "A1")

      assert h_align == "center"
      assert v_align == "middle"
      # Default value
      assert wrap_text == false
      # Default value
      assert rotation == 0

      # Test on a cell that doesn't exist (should return defaults)
      {:ok, default_h_align} =
        UmyaSpreadsheet.get_cell_horizontal_alignment(spreadsheet, "Sheet1", "Z99")

      {:ok, default_v_align} =
        UmyaSpreadsheet.get_cell_vertical_alignment(spreadsheet, "Sheet1", "Z99")

      {:ok, default_wrap} = UmyaSpreadsheet.get_cell_wrap_text(spreadsheet, "Sheet1", "Z99")

      {:ok, default_rotation} =
        UmyaSpreadsheet.get_cell_text_rotation(spreadsheet, "Sheet1", "Z99")

      assert default_h_align == "general"
      assert default_v_align == "bottom"
      assert default_wrap == false
      assert default_rotation == 0
    end

    test "border getter functions return correct values" do
      # Create a new spreadsheet
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Set cell value
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Test Borders")

      # Set border properties using existing setter functions
      UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "B1", "top", "thin")
      UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "B1", "bottom", "medium")
      # Note: There is no set_border_color function, only style
      UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "B1", "left", "dashed")
      UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "B1", "right", "dotted")

      # Test border getters
      {:ok, top_style} = UmyaSpreadsheet.get_border_style(spreadsheet, "Sheet1", "B1", "top")

      {:ok, bottom_style} =
        UmyaSpreadsheet.get_border_style(spreadsheet, "Sheet1", "B1", "bottom")

      {:ok, left_style} = UmyaSpreadsheet.get_border_style(spreadsheet, "Sheet1", "B1", "left")
      {:ok, right_style} = UmyaSpreadsheet.get_border_style(spreadsheet, "Sheet1", "B1", "right")

      {:ok, top_color} = UmyaSpreadsheet.get_border_color(spreadsheet, "Sheet1", "B1", "top")

      {:ok, bottom_color} =
        UmyaSpreadsheet.get_border_color(spreadsheet, "Sheet1", "B1", "bottom")

      {:ok, left_color} = UmyaSpreadsheet.get_border_color(spreadsheet, "Sheet1", "B1", "left")

      assert top_style == "thin"
      assert bottom_style == "medium"
      # Set to dashed
      assert left_style == "dashed"
      # Set to dotted
      assert right_style == "dotted"

      # Default black
      assert top_color == "FF000000"
      # Default black
      assert bottom_color == "FF000000"
      # Default black
      assert left_color == "FF000000"

      # Test invalid border position
      {:ok, invalid_style} =
        UmyaSpreadsheet.get_border_style(spreadsheet, "Sheet1", "B1", "invalid")

      assert invalid_style == "none"
    end

    test "fill/background getter functions return correct values" do
      # Create a new spreadsheet
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Set cell value
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C1", "Test Fill")

      # Set background color using existing setter functions
      UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "C1", "FFFF00")

      # Test fill getters
      {:ok, bg_color} = UmyaSpreadsheet.get_cell_background_color(spreadsheet, "Sheet1", "C1")
      {:ok, fg_color} = UmyaSpreadsheet.get_cell_foreground_color(spreadsheet, "Sheet1", "C1")
      {:ok, pattern_type} = UmyaSpreadsheet.get_cell_pattern_type(spreadsheet, "Sheet1", "C1")

      # Note: might have alpha channel
      assert bg_color == "FFFFFF00"
      # Default black foreground
      assert fg_color == "FF000000"
      # Depending on implementation
      assert pattern_type == "solid" || pattern_type == "none"

      # Test on cell without formatting
      {:ok, default_bg} = UmyaSpreadsheet.get_cell_background_color(spreadsheet, "Sheet1", "D1")
      {:ok, default_fg} = UmyaSpreadsheet.get_cell_foreground_color(spreadsheet, "Sheet1", "D1")
      {:ok, default_pattern} = UmyaSpreadsheet.get_cell_pattern_type(spreadsheet, "Sheet1", "D1")

      # Default white
      assert default_bg == "FFFFFFFF"
      # Default black
      assert default_fg == "FF000000"
      # Default pattern type
      assert default_pattern == "none"
    end

    test "number format getter functions return correct values" do
      # Create a new spreadsheet
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Set cell value and format
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D1", "123.45")
      UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "D1", "0.00")

      # Test number format getters
      {:ok, format_id} = UmyaSpreadsheet.get_cell_number_format_id(spreadsheet, "Sheet1", "D1")
      {:ok, format_code} = UmyaSpreadsheet.get_cell_format_code(spreadsheet, "Sheet1", "D1")

      assert is_integer(format_id)
      assert format_code == "0.00"

      # Test on cell without formatting
      {:ok, default_id} = UmyaSpreadsheet.get_cell_number_format_id(spreadsheet, "Sheet1", "E1")
      {:ok, default_code} = UmyaSpreadsheet.get_cell_format_code(spreadsheet, "Sheet1", "E1")

      assert default_id == 0
      assert default_code == "General"
    end

    test "protection getter functions return correct values" do
      # Create a new spreadsheet
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Set cell value
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "E1", "Test Protection")

      # Test protection getters (note: these might return defaults for now)
      {:ok, locked} = UmyaSpreadsheet.get_cell_locked(spreadsheet, "Sheet1", "E1")
      {:ok, hidden} = UmyaSpreadsheet.get_cell_hidden(spreadsheet, "Sheet1", "E1")

      assert is_boolean(locked)
      assert is_boolean(hidden)

      # These are expected to be defaults since we haven't set protection properties
      # Default locked value
      assert locked == true
      # Default hidden value (implementation limitation noted)
      assert hidden == false
    end

    test "error handling for invalid sheet names" do
      # Create a new spreadsheet
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Test all getter functions with invalid sheet name
      assert {:error, _} =
               UmyaSpreadsheet.get_cell_horizontal_alignment(spreadsheet, "InvalidSheet", "A1")

      assert {:error, _} =
               UmyaSpreadsheet.get_cell_vertical_alignment(spreadsheet, "InvalidSheet", "A1")

      assert {:error, _} = UmyaSpreadsheet.get_cell_wrap_text(spreadsheet, "InvalidSheet", "A1")

      assert {:error, _} =
               UmyaSpreadsheet.get_cell_text_rotation(spreadsheet, "InvalidSheet", "A1")

      assert {:error, _} =
               UmyaSpreadsheet.get_border_style(spreadsheet, "InvalidSheet", "A1", "top")

      assert {:error, _} =
               UmyaSpreadsheet.get_border_color(spreadsheet, "InvalidSheet", "A1", "top")

      assert {:error, _} =
               UmyaSpreadsheet.get_cell_background_color(spreadsheet, "InvalidSheet", "A1")

      assert {:error, _} =
               UmyaSpreadsheet.get_cell_foreground_color(spreadsheet, "InvalidSheet", "A1")

      assert {:error, _} =
               UmyaSpreadsheet.get_cell_pattern_type(spreadsheet, "InvalidSheet", "A1")

      assert {:error, _} =
               UmyaSpreadsheet.get_cell_number_format_id(spreadsheet, "InvalidSheet", "A1")

      assert {:error, _} = UmyaSpreadsheet.get_cell_format_code(spreadsheet, "InvalidSheet", "A1")

      assert {:error, _} = UmyaSpreadsheet.get_cell_locked(spreadsheet, "InvalidSheet", "A1")
      assert {:error, _} = UmyaSpreadsheet.get_cell_hidden(spreadsheet, "InvalidSheet", "A1")
    end

    test "comprehensive getter test with file save/load" do
      # Create a new spreadsheet
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Set various formatting properties
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "F1", "Comprehensive Test")
      UmyaSpreadsheet.set_cell_alignment(spreadsheet, "Sheet1", "F1", "center", "middle")
      UmyaSpreadsheet.set_border_style(spreadsheet, "Sheet1", "F1", "top", "thick")
      UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "F1", "FFFF00")
      # Text format
      UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "F1", "@")

      # Save the file
      UmyaSpreadsheet.write(spreadsheet, @output_path)
      assert File.exists?(@output_path)

      # Read the file back
      {:ok, loaded_spreadsheet} = UmyaSpreadsheet.read(@output_path)

      # Verify all getter functions work on the loaded spreadsheet
      {:ok, h_align} =
        UmyaSpreadsheet.get_cell_horizontal_alignment(loaded_spreadsheet, "Sheet1", "F1")

      {:ok, v_align} =
        UmyaSpreadsheet.get_cell_vertical_alignment(loaded_spreadsheet, "Sheet1", "F1")

      {:ok, border_style} =
        UmyaSpreadsheet.get_border_style(loaded_spreadsheet, "Sheet1", "F1", "top")

      {:ok, bg_color} =
        UmyaSpreadsheet.get_cell_background_color(loaded_spreadsheet, "Sheet1", "F1")

      {:ok, format_code} =
        UmyaSpreadsheet.get_cell_format_code(loaded_spreadsheet, "Sheet1", "F1")

      assert h_align == "center"
      assert v_align == "middle"
      assert border_style == "thick"
      # Might have alpha channel variations
      assert String.contains?(bg_color, "FFFF00")
      assert format_code == "@"

      # Clean up
      File.rm(@output_path)
    end
  end
end
