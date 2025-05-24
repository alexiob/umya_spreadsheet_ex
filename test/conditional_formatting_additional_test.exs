defmodule ConditionalFormattingAdditionalTest do
  use ExUnit.Case
  doctest UmyaSpreadsheet

  alias UmyaSpreadsheet

  describe "Icon Sets Conditional Formatting" do
    test "adds icon set rule with percentile thresholds" do
      # Create a new spreadsheet and add some data
      {:ok, spreadsheet} = UmyaSpreadsheet.new()
      UmyaSpreadsheet.add_sheet(spreadsheet, "IconSets")

      # Add sample data
      for i <- 1..10 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "IconSets", "A#{i}", "#{i * 10}")
      end

      # Add icon set rule with percentile thresholds
      result = UmyaSpreadsheet.add_icon_set(
        spreadsheet,
        "IconSets",
        "A1:A10",
        "3_traffic_lights",
        [
          {"percentile", "33"},
          {"percentile", "67"}
        ]
      )

      assert result == :ok
    end

    test "adds icon set rule with number thresholds" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()
      UmyaSpreadsheet.add_sheet(spreadsheet, "IconSets")

      # Add sample data
      for i <- 1..10 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "IconSets", "A#{i}", "#{i * 5}")
      end

      # Add icon set rule with number thresholds
      result = UmyaSpreadsheet.add_icon_set(
        spreadsheet,
        "IconSets",
        "A1:A10",
        "5_arrows",
        [
          {"number", "10"},
          {"number", "20"},
          {"number", "30"},
          {"number", "40"}
        ]
      )

      assert result == :ok
    end

    test "adds icon set rule with mixed threshold types" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()
      UmyaSpreadsheet.add_sheet(spreadsheet, "IconSets")

      # Add sample data
      for i <- 1..10 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "IconSets", "A#{i}", "#{i * 2}")
      end

      # Add icon set rule with mixed thresholds
      result = UmyaSpreadsheet.add_icon_set(
        spreadsheet,
        "IconSets",
        "A1:A10",
        "3_symbols",
        [
          {"min", ""},
          {"percent", "50"}
        ]
      )

      assert result == :ok
    end

    test "fails with invalid threshold type" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()
      UmyaSpreadsheet.add_sheet(spreadsheet, "IconSets")

      result = UmyaSpreadsheet.add_icon_set(
        spreadsheet,
        "IconSets",
        "A1:A10",
        "3_traffic_lights",
        [
          {"invalid_type", "50"},
          {"percentile", "75"}
        ]
      )

      assert {:error, _} = result
    end

    test "fails with insufficient thresholds" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()
      UmyaSpreadsheet.add_sheet(spreadsheet, "IconSets")

      result = UmyaSpreadsheet.add_icon_set(
        spreadsheet,
        "IconSets",
        "A1:A10",
        "3_traffic_lights",
        [
          {"percentile", "50"}
        ]
      )

      assert {:error, _} = result
    end

    test "fails with too many thresholds" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()
      UmyaSpreadsheet.add_sheet(spreadsheet, "IconSets")

      result = UmyaSpreadsheet.add_icon_set(
        spreadsheet,
        "IconSets",
        "A1:A10",
        "3_traffic_lights",
        [
          {"percentile", "16"},
          {"percentile", "33"},
          {"percentile", "50"},
          {"percentile", "66"},
          {"percentile", "83"},
          {"percentile", "90"}
        ]
      )

      assert {:error, _} = result
    end

    test "fails with nonexistent sheet" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      result = UmyaSpreadsheet.add_icon_set(
        spreadsheet,
        "NonExistentSheet",
        "A1:A10",
        "3_traffic_lights",
        [
          {"percentile", "33"},
          {"percentile", "67"}
        ]
      )

      assert {:error, _} = result
    end

    test "fails with empty cell range" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()
      UmyaSpreadsheet.add_sheet(spreadsheet, "IconSets")

      result = UmyaSpreadsheet.add_icon_set(
        spreadsheet,
        "IconSets",
        "",
        "3_traffic_lights",
        [
          {"percentile", "33"},
          {"percentile", "67"}
        ]
      )

      assert {:error, _} = result
    end
  end

  describe "Above/Below Average Conditional Formatting" do
    test "adds above average rule" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()
      UmyaSpreadsheet.add_sheet(spreadsheet, "AverageTests")

      # Add sample data
      for i <- 1..10 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "AverageTests", "A#{i}", "#{i * 10}")
      end

      # Add above average rule
      result = UmyaSpreadsheet.add_above_below_average_rule(
        spreadsheet,
        "AverageTests",
        "A1:A10",
        "above",
        nil,
        "#00FF00"
      )

      assert result == :ok
    end

    test "adds below average rule" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()
      UmyaSpreadsheet.add_sheet(spreadsheet, "AverageTests")

      # Add sample data
      for i <- 1..10 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "AverageTests", "A#{i}", "#{i * 5}")
      end

      # Add below average rule
      result = UmyaSpreadsheet.add_above_below_average_rule(
        spreadsheet,
        "AverageTests",
        "A1:A10",
        "below",
        nil,
        "#FF0000"
      )

      assert result == :ok
    end

    test "adds above or equal average rule" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()
      UmyaSpreadsheet.add_sheet(spreadsheet, "AverageTests")

      # Add sample data
      for i <- 1..10 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "AverageTests", "A#{i}", "#{i * 2}")
      end

      # Add above or equal average rule
      result = UmyaSpreadsheet.add_above_below_average_rule(
        spreadsheet,
        "AverageTests",
        "A1:A10",
        "above_equal",
        nil,
        "#0000FF"
      )

      assert result == :ok
    end

    test "adds below or equal average rule" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()
      UmyaSpreadsheet.add_sheet(spreadsheet, "AverageTests")

      # Add sample data
      for i <- 1..10 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "AverageTests", "A#{i}", "#{i * 3}")
      end

      # Add below or equal average rule
      result = UmyaSpreadsheet.add_above_below_average_rule(
        spreadsheet,
        "AverageTests",
        "A1:A10",
        "below_equal",
        nil,
        "#FFFF00"
      )

      assert result == :ok
    end

    test "adds above average rule with standard deviation" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()
      UmyaSpreadsheet.add_sheet(spreadsheet, "AverageTests")

      # Add sample data
      for i <- 1..10 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "AverageTests", "A#{i}", "#{i * 8}")
      end

      # Add above average rule with 1 standard deviation
      result = UmyaSpreadsheet.add_above_below_average_rule(
        spreadsheet,
        "AverageTests",
        "A1:A10",
        "above",
        1,
        "#FF00FF"
      )

      assert result == :ok
    end

    test "fails with invalid rule type" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()
      UmyaSpreadsheet.add_sheet(spreadsheet, "AverageTests")

      result = UmyaSpreadsheet.add_above_below_average_rule(
        spreadsheet,
        "AverageTests",
        "A1:A10",
        "invalid_rule",
        nil,
        "#FF0000"
      )

      assert {:error, _} = result
    end

    test "fails with nonexistent sheet" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      result = UmyaSpreadsheet.add_above_below_average_rule(
        spreadsheet,
        "NonExistentSheet",
        "A1:A10",
        "above",
        nil,
        "#FF0000"
      )

      assert {:error, _} = result
    end

    test "fails with empty cell range" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()
      UmyaSpreadsheet.add_sheet(spreadsheet, "AverageTests")

      result = UmyaSpreadsheet.add_above_below_average_rule(
        spreadsheet,
        "AverageTests",
        "",
        "above",
        nil,
        "#FF0000"
      )

      assert {:error, _} = result
    end
  end

  describe "Integration Tests" do
    test "creates comprehensive conditional formatting example with new features" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Sheet 1: Icon Sets
      UmyaSpreadsheet.add_sheet(spreadsheet, "IconSets")
      for i <- 1..20 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "IconSets", "A#{i}", "#{:rand.uniform(100)}")
      end

      # Add 3-icon set
      :ok = UmyaSpreadsheet.add_icon_set(
        spreadsheet,
        "IconSets",
        "A1:A20",
        "3_traffic_lights",
        [
          {"percentile", "33"},
          {"percentile", "67"}
        ]
      )

      # Sheet 2: Average Rules
      UmyaSpreadsheet.add_sheet(spreadsheet, "AverageRules")
      for i <- 1..20 do
        UmyaSpreadsheet.set_cell_value(spreadsheet, "AverageRules", "A#{i}", "#{:rand.uniform(100)}")
      end

      # Add above average rule
      :ok = UmyaSpreadsheet.add_above_below_average_rule(
        spreadsheet,
        "AverageRules",
        "A1:A20",
        "above",
        nil,
        "#00FF00"
      )

      # Add below average rule
      :ok = UmyaSpreadsheet.add_above_below_average_rule(
        spreadsheet,
        "AverageRules",
        "B1:B20",
        "below",
        nil,
        "#FF0000"
      )

      # Save and verify the file can be written
      result = UmyaSpreadsheet.write(spreadsheet, "/tmp/conditional_formatting_additional_test.xlsx")
      assert result == :ok

      # Verify the file exists
      assert File.exists?("/tmp/conditional_formatting_additional_test.xlsx")

      # Clean up
      File.rm("/tmp/conditional_formatting_additional_test.xlsx")
    end
  end
end
