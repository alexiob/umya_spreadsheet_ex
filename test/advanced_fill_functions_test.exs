defmodule UmyaSpreadsheet.AdvancedFillFunctionsTest do
  use ExUnit.Case
  alias UmyaSpreadsheet.AdvancedFillFunctions

  setup do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()
    {:ok, spreadsheet: spreadsheet}
  end

  describe "set_gradient_fill/5" do
    test "sets basic gradient fill with two colors", %{spreadsheet: spreadsheet} do
      gradient_stops = [{0.0, "#FF0000"}, {1.0, "#0000FF"}]

      assert :ok =
               AdvancedFillFunctions.set_gradient_fill(
                 spreadsheet,
                 "Sheet1",
                 "A1",
                 45.0,
                 gradient_stops
               )
    end

    test "sets complex gradient fill with multiple colors", %{spreadsheet: spreadsheet} do
      gradient_stops = [
        # Red
        {0.0, "#FF0000"},
        # Yellow
        {0.25, "#FFFF00"},
        # Green
        {0.5, "#00FF00"},
        # Cyan
        {0.75, "#00FFFF"},
        # Blue
        {1.0, "#0000FF"}
      ]

      assert :ok =
               AdvancedFillFunctions.set_gradient_fill(
                 spreadsheet,
                 "Sheet1",
                 "B1",
                 90.0,
                 gradient_stops
               )
    end

    test "returns error for invalid color format", %{spreadsheet: spreadsheet} do
      gradient_stops = [{0.0, "invalid_color"}, {1.0, "#0000FF"}]

      assert {:error, _reason} =
               AdvancedFillFunctions.set_gradient_fill(
                 spreadsheet,
                 "Sheet1",
                 "A1",
                 45.0,
                 gradient_stops
               )
    end

    test "returns error for non-existent sheet", %{spreadsheet: spreadsheet} do
      gradient_stops = [{0.0, "#FF0000"}, {1.0, "#0000FF"}]

      assert {:error, _reason} =
               AdvancedFillFunctions.set_gradient_fill(
                 spreadsheet,
                 "NonExistentSheet",
                 "A1",
                 45.0,
                 gradient_stops
               )
    end

    test "handles empty gradient stops list", %{spreadsheet: spreadsheet} do
      assert {:error, _reason} =
               AdvancedFillFunctions.set_gradient_fill(
                 spreadsheet,
                 "Sheet1",
                 "A1",
                 45.0,
                 []
               )
    end
  end

  describe "set_linear_gradient_fill/5 and /6" do
    test "sets linear gradient with default angle", %{spreadsheet: spreadsheet} do
      assert :ok =
               AdvancedFillFunctions.set_linear_gradient_fill(
                 spreadsheet,
                 "Sheet1",
                 "C1",
                 "#FF0000",
                 "#0000FF"
               )
    end

    test "sets linear gradient with custom angle", %{spreadsheet: spreadsheet} do
      assert :ok =
               AdvancedFillFunctions.set_linear_gradient_fill(
                 spreadsheet,
                 "Sheet1",
                 "D1",
                 "#00FF00",
                 "#FFFF00",
                 135.0
               )
    end

    test "returns error for invalid colors", %{spreadsheet: spreadsheet} do
      assert {:error, _reason} =
               AdvancedFillFunctions.set_linear_gradient_fill(
                 spreadsheet,
                 "Sheet1",
                 "A1",
                 "invalid",
                 "#0000FF"
               )
    end
  end

  describe "set_radial_gradient_fill/5" do
    test "sets radial gradient from center to edge", %{spreadsheet: spreadsheet} do
      assert :ok =
               AdvancedFillFunctions.set_radial_gradient_fill(
                 spreadsheet,
                 "Sheet1",
                 "E1",
                 "#FFFFFF",
                 "#000000"
               )
    end

    test "returns error for invalid center color", %{spreadsheet: spreadsheet} do
      assert {:error, _reason} =
               AdvancedFillFunctions.set_radial_gradient_fill(
                 spreadsheet,
                 "Sheet1",
                 "A1",
                 "invalid",
                 "#000000"
               )
    end
  end

  describe "set_three_color_gradient_fill/6 and /7" do
    test "sets three color gradient with default angle", %{spreadsheet: spreadsheet} do
      assert :ok =
               AdvancedFillFunctions.set_three_color_gradient_fill(
                 spreadsheet,
                 "Sheet1",
                 "F1",
                 "#FF0000",
                 "#FFFF00",
                 "#00FF00"
               )
    end

    test "sets three color gradient with custom angle", %{spreadsheet: spreadsheet} do
      assert :ok =
               AdvancedFillFunctions.set_three_color_gradient_fill(
                 spreadsheet,
                 "Sheet1",
                 "G1",
                 "#FF0000",
                 "#FFFF00",
                 "#00FF00",
                 45.0
               )
    end

    test "returns error for invalid middle color", %{spreadsheet: spreadsheet} do
      assert {:error, _reason} =
               AdvancedFillFunctions.set_three_color_gradient_fill(
                 spreadsheet,
                 "Sheet1",
                 "A1",
                 "#FF0000",
                 "invalid",
                 "#00FF00"
               )
    end
  end

  describe "set_custom_gradient_fill/5 and /6" do
    test "sets custom gradient with validation enabled", %{spreadsheet: spreadsheet} do
      gradient_stops = [
        {0.2, "#FF0000"},
        {0.8, "#0000FF"}
      ]

      assert :ok =
               AdvancedFillFunctions.set_custom_gradient_fill(
                 spreadsheet,
                 "Sheet1",
                 "H1",
                 60.0,
                 gradient_stops,
                 true
               )
    end

    test "sets custom gradient with validation disabled", %{spreadsheet: spreadsheet} do
      gradient_stops = [
        # Out of order - should work with validation disabled
        {0.8, "#FF0000"},
        {0.2, "#0000FF"}
      ]

      assert :ok =
               AdvancedFillFunctions.set_custom_gradient_fill(
                 spreadsheet,
                 "Sheet1",
                 "I1",
                 60.0,
                 gradient_stops,
                 false
               )
    end

    test "returns error for invalid position values with validation", %{spreadsheet: spreadsheet} do
      gradient_stops = [
        # Invalid position
        {-0.1, "#FF0000"},
        {1.0, "#0000FF"}
      ]

      assert {:error, _reason} =
               AdvancedFillFunctions.set_custom_gradient_fill(
                 spreadsheet,
                 "Sheet1",
                 "A1",
                 60.0,
                 gradient_stops,
                 true
               )
    end
  end

  describe "get_gradient_fill/3" do
    test "retrieves gradient fill after setting it", %{spreadsheet: spreadsheet} do
      gradient_stops = [{0.0, "#FF0000"}, {1.0, "#0000FF"}]

      :ok =
        AdvancedFillFunctions.set_gradient_fill(
          spreadsheet,
          "Sheet1",
          "J1",
          45.0,
          gradient_stops
        )

      assert {:ok, {degree, stops}} =
               AdvancedFillFunctions.get_gradient_fill(
                 spreadsheet,
                 "Sheet1",
                 "J1"
               )

      assert degree == 45.0
      assert is_list(stops)
      assert length(stops) == 2
    end

    test "returns error for cell without gradient fill", %{spreadsheet: spreadsheet} do
      assert {:error, _reason} =
               AdvancedFillFunctions.get_gradient_fill(
                 spreadsheet,
                 "Sheet1",
                 "K1"
               )
    end

    test "returns error for non-existent sheet", %{spreadsheet: spreadsheet} do
      assert {:error, _reason} =
               AdvancedFillFunctions.get_gradient_fill(
                 spreadsheet,
                 "NonExistentSheet",
                 "A1"
               )
    end
  end

  describe "set_pattern_fill/4, /5, and /6" do
    test "sets solid pattern fill", %{spreadsheet: spreadsheet} do
      assert :ok =
               AdvancedFillFunctions.set_pattern_fill(
                 spreadsheet,
                 "Sheet1",
                 "L1",
                 :solid,
                 "#FF0000"
               )
    end

    test "sets pattern fill with background color", %{spreadsheet: spreadsheet} do
      assert :ok =
               AdvancedFillFunctions.set_pattern_fill(
                 spreadsheet,
                 "Sheet1",
                 "M1",
                 :dark_grid,
                 "#FF0000",
                 "#FFFFFF"
               )
    end

    test "sets pattern fill using string pattern type", %{spreadsheet: spreadsheet} do
      assert :ok =
               AdvancedFillFunctions.set_pattern_fill(
                 spreadsheet,
                 "Sheet1",
                 "N1",
                 "light_horizontal",
                 "#00FF00"
               )
    end

    test "sets various pattern types", %{spreadsheet: spreadsheet} do
      patterns = [
        :solid,
        :dark_gray,
        :medium_gray,
        :light_gray,
        :gray_125,
        :gray_0625,
        :dark_horizontal,
        :dark_vertical,
        :dark_down,
        :dark_up,
        :dark_grid,
        :dark_trellis,
        :light_horizontal,
        :light_vertical,
        :light_down,
        :light_up,
        :light_grid,
        :light_trellis
      ]

      Enum.with_index(patterns, 1)
      |> Enum.each(fn {pattern, index} ->
        cell = "A#{index + 1}"

        assert :ok =
                 AdvancedFillFunctions.set_pattern_fill(
                   spreadsheet,
                   "Sheet1",
                   cell,
                   pattern,
                   "#FF0000"
                 )
      end)
    end

    test "returns error for invalid pattern type", %{spreadsheet: spreadsheet} do
      assert {:error, _reason} =
               AdvancedFillFunctions.set_pattern_fill(
                 spreadsheet,
                 "Sheet1",
                 "A1",
                 :invalid_pattern,
                 "#FF0000"
               )
    end

    test "returns error for invalid foreground color", %{spreadsheet: spreadsheet} do
      assert {:error, _reason} =
               AdvancedFillFunctions.set_pattern_fill(
                 spreadsheet,
                 "Sheet1",
                 "A1",
                 :solid,
                 "invalid_color"
               )
    end
  end

  describe "get_pattern_fill/3" do
    test "retrieves pattern fill after setting it", %{spreadsheet: spreadsheet} do
      :ok =
        AdvancedFillFunctions.set_pattern_fill(
          spreadsheet,
          "Sheet1",
          "O1",
          :dark_grid,
          "#FF0000",
          "#FFFFFF"
        )

      assert {:ok, {pattern, fg_color, bg_color}} =
               AdvancedFillFunctions.get_pattern_fill(
                 spreadsheet,
                 "Sheet1",
                 "O1"
               )

      assert pattern == "dark_grid"
      assert is_binary(fg_color)
      assert is_binary(bg_color)
    end

    test "returns error for cell without pattern fill", %{spreadsheet: spreadsheet} do
      assert {:error, _reason} =
               AdvancedFillFunctions.get_pattern_fill(
                 spreadsheet,
                 "Sheet1",
                 "P1"
               )
    end

    test "returns error for non-existent sheet", %{spreadsheet: spreadsheet} do
      assert {:error, _reason} =
               AdvancedFillFunctions.get_pattern_fill(
                 spreadsheet,
                 "NonExistentSheet",
                 "A1"
               )
    end
  end

  describe "clear_fill/3" do
    test "clears gradient fill", %{spreadsheet: spreadsheet} do
      # Set a gradient fill
      gradient_stops = [{0.0, "#FF0000"}, {1.0, "#0000FF"}]

      :ok =
        AdvancedFillFunctions.set_gradient_fill(
          spreadsheet,
          "Sheet1",
          "Q1",
          45.0,
          gradient_stops
        )

      # Verify it exists
      assert {:ok, _} = AdvancedFillFunctions.get_gradient_fill(spreadsheet, "Sheet1", "Q1")

      # Clear the fill
      assert :ok = AdvancedFillFunctions.clear_fill(spreadsheet, "Sheet1", "Q1")

      # Verify it's gone
      assert {:error, _reason} =
               AdvancedFillFunctions.get_gradient_fill(spreadsheet, "Sheet1", "Q1")
    end

    test "clears pattern fill", %{spreadsheet: spreadsheet} do
      # Set a pattern fill
      :ok =
        AdvancedFillFunctions.set_pattern_fill(
          spreadsheet,
          "Sheet1",
          "R1",
          :solid,
          "#FF0000"
        )

      # Verify it exists
      assert {:ok, _} = AdvancedFillFunctions.get_pattern_fill(spreadsheet, "Sheet1", "R1")

      # Clear the fill
      assert :ok = AdvancedFillFunctions.clear_fill(spreadsheet, "Sheet1", "R1")

      # Verify it's gone
      assert {:error, _reason} =
               AdvancedFillFunctions.get_pattern_fill(spreadsheet, "Sheet1", "R1")
    end

    test "works on cells without existing fill", %{spreadsheet: spreadsheet} do
      assert :ok = AdvancedFillFunctions.clear_fill(spreadsheet, "Sheet1", "S1")
    end

    test "returns error for non-existent sheet", %{spreadsheet: spreadsheet} do
      assert {:error, _reason} =
               AdvancedFillFunctions.clear_fill(
                 spreadsheet,
                 "NonExistentSheet",
                 "A1"
               )
    end
  end

  describe "integration tests" do
    test "can apply multiple different fills to different cells", %{spreadsheet: spreadsheet} do
      # Set gradient fill on A1
      gradient_stops = [{0.0, "#FF0000"}, {1.0, "#0000FF"}]

      :ok =
        AdvancedFillFunctions.set_gradient_fill(
          spreadsheet,
          "Sheet1",
          "A1",
          45.0,
          gradient_stops
        )

      # Set pattern fill on B1
      :ok =
        AdvancedFillFunctions.set_pattern_fill(
          spreadsheet,
          "Sheet1",
          "B1",
          :dark_grid,
          "#00FF00",
          "#FFFFFF"
        )

      # Set linear gradient on C1
      :ok =
        AdvancedFillFunctions.set_linear_gradient_fill(
          spreadsheet,
          "Sheet1",
          "C1",
          "#FFFF00",
          "#FF00FF",
          90.0
        )

      # Verify all fills exist
      assert {:ok, _} = AdvancedFillFunctions.get_gradient_fill(spreadsheet, "Sheet1", "A1")
      assert {:ok, _} = AdvancedFillFunctions.get_pattern_fill(spreadsheet, "Sheet1", "B1")
      assert {:ok, _} = AdvancedFillFunctions.get_gradient_fill(spreadsheet, "Sheet1", "C1")
    end

    test "can overwrite existing fills", %{spreadsheet: spreadsheet} do
      # Set initial gradient fill
      gradient_stops1 = [{0.0, "#FF0000"}, {1.0, "#0000FF"}]

      :ok =
        AdvancedFillFunctions.set_gradient_fill(
          spreadsheet,
          "Sheet1",
          "A1",
          45.0,
          gradient_stops1
        )

      # Overwrite with pattern fill
      :ok =
        AdvancedFillFunctions.set_pattern_fill(
          spreadsheet,
          "Sheet1",
          "A1",
          :solid,
          "#00FF00"
        )

      # Gradient should be gone, pattern should exist
      assert {:error, _} = AdvancedFillFunctions.get_gradient_fill(spreadsheet, "Sheet1", "A1")
      assert {:ok, _} = AdvancedFillFunctions.get_pattern_fill(spreadsheet, "Sheet1", "A1")

      # Overwrite with new gradient
      gradient_stops2 = [{0.0, "#FFFF00"}, {1.0, "#FF00FF"}]

      :ok =
        AdvancedFillFunctions.set_gradient_fill(
          spreadsheet,
          "Sheet1",
          "A1",
          90.0,
          gradient_stops2
        )

      # Pattern should be gone, gradient should exist
      assert {:error, _} = AdvancedFillFunctions.get_pattern_fill(spreadsheet, "Sheet1", "A1")

      assert {:ok, {degree, _}} =
               AdvancedFillFunctions.get_gradient_fill(spreadsheet, "Sheet1", "A1")

      assert degree == 90.0
    end

    test "can work with ranges of cells", %{spreadsheet: spreadsheet} do
      # Apply same fill to multiple cells
      cells = ["A1", "A2", "A3", "B1", "B2", "B3"]

      Enum.each(cells, fn cell ->
        :ok =
          AdvancedFillFunctions.set_linear_gradient_fill(
            spreadsheet,
            "Sheet1",
            cell,
            "#FF0000",
            "#0000FF",
            45.0
          )
      end)

      # Verify all cells have the fill
      Enum.each(cells, fn cell ->
        assert {:ok, {degree, _}} =
                 AdvancedFillFunctions.get_gradient_fill(
                   spreadsheet,
                   "Sheet1",
                   cell
                 )

        assert degree == 45.0
      end)
    end
  end

  describe "color format validation" do
    test "accepts various valid color formats", %{spreadsheet: spreadsheet} do
      valid_colors = [
        # 6-digit hex
        "#FF0000",
        # 8-digit hex with alpha
        "#FFFF0000",
        # 6-digit without #
        "FF0000",
        # 8-digit without #
        "FFFF0000"
      ]

      Enum.with_index(valid_colors, 1)
      |> Enum.each(fn {color, index} ->
        cell = "A#{index}"

        assert :ok =
                 AdvancedFillFunctions.set_pattern_fill(
                   spreadsheet,
                   "Sheet1",
                   cell,
                   :solid,
                   color
                 )
      end)
    end

    test "rejects invalid color formats", %{spreadsheet: spreadsheet} do
      invalid_colors = [
        # Invalid hex characters
        "#GG0000",
        # Too short
        "#FF00",
        # Too long
        "123456789",
        # Empty string
        "",
        # Just hash
        "#"
      ]

      Enum.each(invalid_colors, fn color ->
        assert {:error, _} =
                 AdvancedFillFunctions.set_pattern_fill(
                   spreadsheet,
                   "Sheet1",
                   "A1",
                   :solid,
                   color
                 )
      end)
    end
  end
end
