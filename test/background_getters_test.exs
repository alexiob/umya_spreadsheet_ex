defmodule UmyaSpreadsheet.BackgroundGettersTest do
  use ExUnit.Case
  alias UmyaSpreadsheet.BackgroundFunctions

  setup do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()
    {:ok, spreadsheet: spreadsheet}
  end

  describe "get_cell_background_color/3" do
    test "retrieves background color after setting it", %{spreadsheet: spreadsheet} do
      # Set background color
      :ok = BackgroundFunctions.set_background_color(spreadsheet, "Sheet1", "A1", "#FFFF00")

      # Get background color
      {:ok, color} = BackgroundFunctions.get_cell_background_color(spreadsheet, "Sheet1", "A1")

      # Should return the color in hex format
      assert is_binary(color)
      # ARGB format
      assert String.length(color) == 8
    end

    test "returns default color for cell without background", %{spreadsheet: spreadsheet} do
      {:ok, color} = BackgroundFunctions.get_cell_background_color(spreadsheet, "Sheet1", "B1")

      # Should return some default color
      assert is_binary(color)
    end

    test "returns error for non-existent sheet", %{spreadsheet: spreadsheet} do
      {:error, _reason} =
        BackgroundFunctions.get_cell_background_color(spreadsheet, "NonExistentSheet", "A1")
    end
  end

  describe "get_cell_foreground_color/3" do
    test "retrieves foreground color", %{spreadsheet: spreadsheet} do
      {:ok, color} = BackgroundFunctions.get_cell_foreground_color(spreadsheet, "Sheet1", "A1")

      # Should return a color value
      assert is_binary(color)
      # ARGB format
      assert String.length(color) == 8
    end

    test "returns error for non-existent sheet", %{spreadsheet: spreadsheet} do
      {:error, _reason} =
        BackgroundFunctions.get_cell_foreground_color(spreadsheet, "NonExistentSheet", "A1")
    end
  end

  describe "get_cell_pattern_type/3" do
    test "retrieves pattern type", %{spreadsheet: spreadsheet} do
      {:ok, pattern} = BackgroundFunctions.get_cell_pattern_type(spreadsheet, "Sheet1", "A1")

      # Should return a pattern type string
      assert is_binary(pattern)
      # Common pattern types: "none", "solid", "mediumGray", etc.
      assert pattern in [
               "none",
               "solid",
               "mediumGray",
               "lightGray",
               "darkGray",
               "darkHorizontal",
               "darkVertical",
               "darkDown",
               "darkUp",
               "darkGrid",
               "darkTrellis",
               "lightHorizontal",
               "lightVertical",
               "lightDown",
               "lightUp",
               "lightGrid",
               "lightTrellis",
               "gray125",
               "gray0625"
             ]
    end

    test "returns error for non-existent sheet", %{spreadsheet: spreadsheet} do
      {:error, _reason} =
        BackgroundFunctions.get_cell_pattern_type(spreadsheet, "NonExistentSheet", "A1")
    end
  end

  describe "integration with background setters" do
    test "getters work consistently with setters", %{spreadsheet: spreadsheet} do
      # Set a background color
      :ok = BackgroundFunctions.set_background_color(spreadsheet, "Sheet1", "C3", "#FF0000")

      # Verify we can retrieve information about it
      {:ok, bg_color} = BackgroundFunctions.get_cell_background_color(spreadsheet, "Sheet1", "C3")
      {:ok, fg_color} = BackgroundFunctions.get_cell_foreground_color(spreadsheet, "Sheet1", "C3")
      {:ok, pattern} = BackgroundFunctions.get_cell_pattern_type(spreadsheet, "Sheet1", "C3")

      # All should return valid values
      assert is_binary(bg_color)
      assert is_binary(fg_color)
      assert is_binary(pattern)

      # Background color should reflect what we set (may be in different format)
      assert String.length(bg_color) == 8
    end

    test "can inspect multiple cells with different backgrounds", %{spreadsheet: spreadsheet} do
      # Set different background colors
      # Red
      :ok = BackgroundFunctions.set_background_color(spreadsheet, "Sheet1", "A1", "#FF0000")
      # Green
      :ok = BackgroundFunctions.set_background_color(spreadsheet, "Sheet1", "A2", "#00FF00")
      # Blue
      :ok = BackgroundFunctions.set_background_color(spreadsheet, "Sheet1", "A3", "#0000FF")

      # Get all background colors
      {:ok, color1} = BackgroundFunctions.get_cell_background_color(spreadsheet, "Sheet1", "A1")
      {:ok, color2} = BackgroundFunctions.get_cell_background_color(spreadsheet, "Sheet1", "A2")
      {:ok, color3} = BackgroundFunctions.get_cell_background_color(spreadsheet, "Sheet1", "A3")

      # All should be different (though exact format may vary)
      assert color1 != color2
      assert color2 != color3
      assert color1 != color3
    end
  end
end
