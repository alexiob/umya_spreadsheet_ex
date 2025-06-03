defmodule UmyaSpreadsheet.AdvancedConditionalFormattingTest do
  @moduledoc """
  Comprehensive tests for Advanced Conditional Formatting features:
  - ColorScale (2-color and 3-color scales)
  - DataBar (with various min/max configurations)
  - IconSet (with different threshold types and icon styles)
  """

  use ExUnit.Case, async: true

  alias UmyaSpreadsheet

  @output_path "test/result_files/advanced_conditional_formatting"

  setup_all do
    # Make sure the result files directory exists
    File.mkdir_p!("test/result_files")
    :ok
  end

  describe "ColorScale Conditional Formatting" do
    test "creates 2-color scale with default min/max" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add test data
      for i <- 1..10 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A#{i}", "#{i * 10}")
      end

      # Add 2-color scale (white to red)
      assert :ok =
               UmyaSpreadsheet.add_color_scale(
                 spreadsheet,
                 "Sheet1",
                 "A1:A10",
                 "min",
                 nil,
                 # White
                 %{argb: "FFFFFFFF"},
                 "max",
                 nil,
                 # Red
                 %{argb: "FFFF0000"}
               )

      # Save and verify
      assert :ok = UmyaSpreadsheet.write(spreadsheet, "#{@output_path}_2color_default.xlsx")
      assert File.exists?("#{@output_path}_2color_default.xlsx")
    end

    test "creates 2-color scale with custom values" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add test data
      for i <- 1..20 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B#{i}", "#{:rand.uniform(100)}")
      end

      # Add 2-color scale with specific values
      assert :ok =
               UmyaSpreadsheet.add_color_scale(
                 spreadsheet,
                 "Sheet1",
                 "B1:B20",
                 "number",
                 "25",
                 # Blue
                 %{argb: "FF0000FF"},
                 "number",
                 "75",
                 # Red
                 %{argb: "FFFF0000"}
               )

      # Save and verify
      assert :ok = UmyaSpreadsheet.write(spreadsheet, "#{@output_path}_2color_custom.xlsx")
      assert File.exists?("#{@output_path}_2color_custom.xlsx")
    end

    test "creates 3-color scale with percentile thresholds" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add test data
      for i <- 1..15 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C#{i}", "#{i * 5}")
      end

      # Add 3-color scale (red-yellow-green)
      assert :ok =
               UmyaSpreadsheet.add_color_scale(
                 spreadsheet,
                 "Sheet1",
                 "C1:C15",
                 "min",
                 nil,
                 # Red
                 %{argb: "FFFF0000"},
                 "percentile",
                 "50",
                 # Yellow
                 %{argb: "FFFFFF00"},
                 "max",
                 nil,
                 # Green
                 %{argb: "FF00FF00"}
               )

      # Save and verify
      assert :ok = UmyaSpreadsheet.write(spreadsheet, "#{@output_path}_3color_percentile.xlsx")
      assert File.exists?("#{@output_path}_3color_percentile.xlsx")
    end

    test "creates 3-color scale with percent thresholds" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add test data
      for i <- 1..12 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D#{i}", "#{i * 8}")
      end

      # Add 3-color scale with percent thresholds
      assert :ok =
               UmyaSpreadsheet.add_color_scale(
                 spreadsheet,
                 "Sheet1",
                 "D1:D12",
                 "percent",
                 "10",
                 # Purple
                 %{argb: "FF800080"},
                 "percent",
                 "50",
                 # Orange
                 %{argb: "FFFFA500"},
                 "percent",
                 "90",
                 # Dark Green
                 %{argb: "FF008000"}
               )

      # Save and verify
      assert :ok = UmyaSpreadsheet.write(spreadsheet, "#{@output_path}_3color_percent.xlsx")
      assert File.exists?("#{@output_path}_3color_percent.xlsx")
    end

    test "handles invalid color scale parameters gracefully" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add some test data
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "10")

      # Test with empty range
      assert {:error, _} =
               UmyaSpreadsheet.add_color_scale(
                 spreadsheet,
                 "Sheet1",
                 "",
                 "min",
                 nil,
                 %{argb: "FFFFFFFF"},
                 "max",
                 nil,
                 %{argb: "FFFF0000"}
               )
    end
  end

  describe "DataBar Conditional Formatting" do
    test "creates data bar with default min/max" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add test data
      for i <- 1..10 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "E#{i}", "#{i * 15}")
      end

      # Add data bar with default settings
      assert :ok =
               UmyaSpreadsheet.add_data_bar(
                 spreadsheet,
                 "Sheet1",
                 "E1:E10",
                 nil,
                 nil,
                 # Blue
                 "#3366CC"
               )

      # Save and verify
      assert :ok = UmyaSpreadsheet.write(spreadsheet, "#{@output_path}_databar_default.xlsx")
      assert File.exists?("#{@output_path}_databar_default.xlsx")
    end

    test "creates data bar with custom min/max values" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add test data
      for i <- 1..15 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "F#{i}", "#{:rand.uniform(200)}")
      end

      # Add data bar with custom min/max
      assert :ok =
               UmyaSpreadsheet.add_data_bar(
                 spreadsheet,
                 "Sheet1",
                 "F1:F15",
                 {"number", "50"},
                 {"number", "150"},
                 # Orange
                 "#FF6600"
               )

      # Save and verify
      assert :ok = UmyaSpreadsheet.write(spreadsheet, "#{@output_path}_databar_custom.xlsx")
      assert File.exists?("#{@output_path}_databar_custom.xlsx")
    end

    test "creates data bar with percentage thresholds" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add test data
      for i <- 1..12 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "G#{i}", "#{i * 12}")
      end

      # Add data bar with percentage thresholds
      assert :ok =
               UmyaSpreadsheet.add_data_bar(
                 spreadsheet,
                 "Sheet1",
                 "G1:G12",
                 {"percent", "20"},
                 {"percent", "80"},
                 # Green
                 "#009900"
               )

      # Save and verify
      assert :ok = UmyaSpreadsheet.write(spreadsheet, "#{@output_path}_databar_percent.xlsx")
      assert File.exists?("#{@output_path}_databar_percent.xlsx")
    end

    test "creates data bar with percentile thresholds" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add test data
      for i <- 1..20 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "H#{i}", "#{:rand.uniform(300)}")
      end

      # Add data bar with percentile thresholds
      assert :ok =
               UmyaSpreadsheet.add_data_bar(
                 spreadsheet,
                 "Sheet1",
                 "H1:H20",
                 {"percentile", "25"},
                 {"percentile", "75"},
                 # Magenta
                 "#CC0099"
               )

      # Save and verify
      assert :ok = UmyaSpreadsheet.write(spreadsheet, "#{@output_path}_databar_percentile.xlsx")
      assert File.exists?("#{@output_path}_databar_percentile.xlsx")
    end

    test "handles invalid data bar parameters gracefully" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add some test data
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "10")

      # Test with empty color
      assert {:error, _} =
               UmyaSpreadsheet.add_data_bar(
                 spreadsheet,
                 "Sheet1",
                 "A1:A1",
                 nil,
                 nil,
                 ""
               )

      # Test with empty range
      assert {:error, _} =
               UmyaSpreadsheet.add_data_bar(
                 spreadsheet,
                 "Sheet1",
                 "",
                 nil,
                 nil,
                 "#FF0000"
               )
    end
  end

  describe "IconSet Conditional Formatting" do
    test "creates 3-icon set with percentile thresholds" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add test data
      for i <- 1..15 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "I#{i}", "#{i * 6}")
      end

      # Add 3-icon set with percentile thresholds
      assert :ok =
               UmyaSpreadsheet.add_icon_set(
                 spreadsheet,
                 "Sheet1",
                 "I1:I15",
                 "3_traffic_lights",
                 [
                   {"percentile", "33"},
                   {"percentile", "67"}
                 ]
               )

      # Save and verify
      assert :ok = UmyaSpreadsheet.write(spreadsheet, "#{@output_path}_iconset_3_percentile.xlsx")
      assert File.exists?("#{@output_path}_iconset_3_percentile.xlsx")
    end

    test "creates 4-icon set with number thresholds" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add test data
      for i <- 1..20 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "J#{i}", "#{i * 4}")
      end

      # Add 4-icon set with number thresholds
      assert :ok =
               UmyaSpreadsheet.add_icon_set(
                 spreadsheet,
                 "Sheet1",
                 "J1:J20",
                 "4_arrows",
                 [
                   {"number", "20"},
                   {"number", "40"},
                   {"number", "60"}
                 ]
               )

      # Save and verify
      assert :ok = UmyaSpreadsheet.write(spreadsheet, "#{@output_path}_iconset_4_number.xlsx")
      assert File.exists?("#{@output_path}_iconset_4_number.xlsx")
    end

    test "creates 5-icon set with mixed thresholds" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add test data
      for i <- 1..25 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "K#{i}", "#{:rand.uniform(100)}")
      end

      # Add 5-icon set with mixed threshold types
      assert :ok =
               UmyaSpreadsheet.add_icon_set(
                 spreadsheet,
                 "Sheet1",
                 "K1:K25",
                 "5_arrows",
                 [
                   {"percentile", "20"},
                   {"percent", "40"},
                   {"number", "60"},
                   {"percentile", "80"}
                 ]
               )

      # Save and verify
      assert :ok = UmyaSpreadsheet.write(spreadsheet, "#{@output_path}_iconset_5_mixed.xlsx")
      assert File.exists?("#{@output_path}_iconset_5_mixed.xlsx")
    end

    test "creates icon set with min/max thresholds" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add test data
      for i <- 1..10 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "L#{i}", "#{i * 11}")
      end

      # Add icon set with min/max
      assert :ok =
               UmyaSpreadsheet.add_icon_set(
                 spreadsheet,
                 "Sheet1",
                 "L1:L10",
                 "3_symbols",
                 [
                   {"min", ""},
                   {"max", ""}
                 ]
               )

      # Save and verify
      assert :ok = UmyaSpreadsheet.write(spreadsheet, "#{@output_path}_iconset_min_max.xlsx")
      assert File.exists?("#{@output_path}_iconset_min_max.xlsx")
    end

    test "handles invalid icon set parameters gracefully" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add some test data
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "10")

      # Test with invalid threshold type
      assert {:error, _} =
               UmyaSpreadsheet.add_icon_set(
                 spreadsheet,
                 "Sheet1",
                 "A1:A1",
                 "3_traffic_lights",
                 [
                   {"invalid_type", "50"},
                   {"percentile", "75"}
                 ]
               )

      # Test with too few thresholds
      assert {:error, _} =
               UmyaSpreadsheet.add_icon_set(
                 spreadsheet,
                 "Sheet1",
                 "A1:A1",
                 "3_traffic_lights",
                 [
                   {"percentile", "50"}
                 ]
               )

      # Test with too many thresholds
      assert {:error, _} =
               UmyaSpreadsheet.add_icon_set(
                 spreadsheet,
                 "Sheet1",
                 "A1:A1",
                 "3_traffic_lights",
                 [
                   {"percentile", "10"},
                   {"percentile", "20"},
                   {"percentile", "30"},
                   {"percentile", "40"},
                   {"percentile", "50"},
                   {"percentile", "60"}
                 ]
               )

      # Test with empty range
      assert {:error, _} =
               UmyaSpreadsheet.add_icon_set(
                 spreadsheet,
                 "Sheet1",
                 "",
                 "3_traffic_lights",
                 [
                   {"percentile", "33"},
                   {"percentile", "67"}
                 ]
               )
    end
  end

  describe "Getter Functions" do
    test "retrieves color scales from worksheet" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add test data and color scale
      for i <- 1..5 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "M#{i}", "#{i * 20}")
      end

      UmyaSpreadsheet.add_color_scale(
        spreadsheet,
        "Sheet1",
        "M1:M5",
        "min",
        nil,
        %{argb: "FFFFFFFF"},
        "max",
        nil,
        %{argb: "FFFF0000"}
      )

      # Retrieve color scales
      {:ok, color_scales} = UmyaSpreadsheet.get_color_scales(spreadsheet, "Sheet1", nil)
      assert is_list(color_scales)
      assert length(color_scales) >= 1
    end

    test "retrieves data bars from worksheet" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add test data and data bar
      for i <- 1..5 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "N#{i}", "#{i * 25}")
      end

      UmyaSpreadsheet.add_data_bar(
        spreadsheet,
        "Sheet1",
        "N1:N5",
        nil,
        nil,
        "#3366CC"
      )

      # Retrieve data bars
      {:ok, data_bars} = UmyaSpreadsheet.get_data_bars(spreadsheet, "Sheet1", nil)
      assert is_list(data_bars)
      assert length(data_bars) >= 1
    end

    test "retrieves icon sets from worksheet" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add test data and icon set
      for i <- 1..5 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "O#{i}", "#{i * 30}")
      end

      UmyaSpreadsheet.add_icon_set(
        spreadsheet,
        "Sheet1",
        "O1:O5",
        "3_traffic_lights",
        [
          {"percentile", "33"},
          {"percentile", "67"}
        ]
      )

      # Retrieve icon sets
      {:ok, icon_sets} = UmyaSpreadsheet.get_icon_sets(spreadsheet, "Sheet1", nil)
      assert is_list(icon_sets)
      assert length(icon_sets) >= 1
    end

    test "retrieves conditional formatting with range filter" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add test data
      for i <- 1..10 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "P#{i}", "#{i * 7}")
      end

      # Add multiple formatting rules
      UmyaSpreadsheet.add_color_scale(
        spreadsheet,
        "Sheet1",
        "P1:P5",
        "min",
        nil,
        %{argb: "FFFFFFFF"},
        "max",
        nil,
        %{argb: "FFFF0000"}
      )

      UmyaSpreadsheet.add_data_bar(
        spreadsheet,
        "Sheet1",
        "P6:P10",
        nil,
        nil,
        "#00FF00"
      )

      # Test range-specific retrieval
      {:ok, color_scales} = UmyaSpreadsheet.get_color_scales(spreadsheet, "Sheet1", "P1:P5")
      {:ok, data_bars} = UmyaSpreadsheet.get_data_bars(spreadsheet, "Sheet1", "P6:P10")

      assert is_list(color_scales)
      assert is_list(data_bars)
    end
  end

  describe "Complex Scenarios" do
    test "combines multiple advanced formatting types in one worksheet" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Create comprehensive test data
      UmyaSpreadsheet.add_sheet(spreadsheet, "Advanced")

      # Sales data for color scale
      sales_data = [120, 85, 200, 67, 145, 89, 176, 203, 91, 158]

      for {value, i} <- Enum.with_index(sales_data, 1) do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "Advanced", "A#{i}", "#{value}")
      end

      # Performance percentages for data bars
      performance_data = [95, 78, 88, 92, 85, 91, 87, 94, 89, 96]

      for {value, i} <- Enum.with_index(performance_data, 1) do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "Advanced", "B#{i}", "#{value}")
      end

      # Rating scores for icon sets
      rating_data = [4.5, 3.2, 4.8, 2.9, 4.1, 3.7, 4.6, 4.9, 3.5, 4.3]

      for {value, i} <- Enum.with_index(rating_data, 1) do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "Advanced", "C#{i}", "#{value}")
      end

      # Apply different formatting to each column
      # Sales: 3-color scale (red-yellow-green)
      assert :ok =
               UmyaSpreadsheet.add_color_scale(
                 spreadsheet,
                 "Advanced",
                 "A1:A10",
                 "min",
                 nil,
                 # Red
                 %{argb: "FFFF0000"},
                 "percentile",
                 "50",
                 # Yellow
                 %{argb: "FFFFFF00"},
                 "max",
                 nil,
                 # Green
                 %{argb: "FF00FF00"}
               )

      # Performance: Data bars
      assert :ok =
               UmyaSpreadsheet.add_data_bar(
                 spreadsheet,
                 "Advanced",
                 "B1:B10",
                 {"number", "75"},
                 {"number", "100"},
                 "#0066CC"
               )

      # Ratings: 5-icon set
      assert :ok =
               UmyaSpreadsheet.add_icon_set(
                 spreadsheet,
                 "Advanced",
                 "C1:C10",
                 "5_quarters",
                 [
                   {"number", "2.5"},
                   {"number", "3.0"},
                   {"number", "3.5"},
                   {"number", "4.0"}
                 ]
               )

      # Save and verify
      assert :ok = UmyaSpreadsheet.write(spreadsheet, "#{@output_path}_complex_scenario.xlsx")
      assert File.exists?("#{@output_path}_complex_scenario.xlsx")

      # Verify all formatting rules exist
      {:ok, color_scales} = UmyaSpreadsheet.get_color_scales(spreadsheet, "Advanced", nil)
      {:ok, data_bars} = UmyaSpreadsheet.get_data_bars(spreadsheet, "Advanced", nil)
      {:ok, icon_sets} = UmyaSpreadsheet.get_icon_sets(spreadsheet, "Advanced", nil)

      assert length(color_scales) >= 1
      assert length(data_bars) >= 1
      assert length(icon_sets) >= 1
    end

    test "stress test with large ranges" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Create large dataset
      UmyaSpreadsheet.add_sheet(spreadsheet, "Stress")

      # Generate 100 rows of test data
      for i <- 1..100 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "Stress", "A#{i}", "#{:rand.uniform(1000)}")
      end

      # Apply formatting to large range
      assert :ok =
               UmyaSpreadsheet.add_color_scale(
                 spreadsheet,
                 "Stress",
                 "A1:A100",
                 "min",
                 nil,
                 %{argb: "FFFFFFFF"},
                 "percentile",
                 "50",
                 %{argb: "FFFFFF00"},
                 "max",
                 nil,
                 %{argb: "FFFF0000"}
               )

      # Save and verify
      assert :ok = UmyaSpreadsheet.write(spreadsheet, "#{@output_path}_stress_test.xlsx")
      assert File.exists?("#{@output_path}_stress_test.xlsx")
    end
  end

  describe "Integration with Existing Features" do
    test "combines advanced formatting with basic conditional formatting" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add test data
      for i <- 1..10 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "Q#{i}", "#{i * 13}")
        UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "R#{i}", "#{i * 17}")
      end

      # Add basic cell value rule
      assert :ok =
               UmyaSpreadsheet.add_cell_value_rule(
                 spreadsheet,
                 "Sheet1",
                 "Q1:Q10",
                 "greaterThan",
                 "65",
                 nil,
                 # Gold
                 "#FFD700"
               )

      # Add advanced color scale
      assert :ok =
               UmyaSpreadsheet.add_color_scale(
                 spreadsheet,
                 "Sheet1",
                 "R1:R10",
                 "min",
                 nil,
                 # Sky Blue
                 %{argb: "FF87CEEB"},
                 "max",
                 nil,
                 # Royal Blue
                 %{argb: "FF4169E1"}
               )

      # Save and verify
      assert :ok = UmyaSpreadsheet.write(spreadsheet, "#{@output_path}_integration.xlsx")
      assert File.exists?("#{@output_path}_integration.xlsx")
    end
  end
end
