defmodule UmyaSpreadsheet.ConditionalFormattingGetterTest do
  use ExUnit.Case, async: true

  alias UmyaSpreadsheet
  alias UmyaSpreadsheet.ConditionalFormatting

  @output_path "test/result_files/conditional_formatting_getter_test.xlsx"

  setup_all do
    # Make sure the result files directory exists
    File.mkdir_p!("test/result_files")
    :ok
  end

  test "get conditional formatting rules" do
    # Create a new spreadsheet with different conditional formatting rules
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add test data
    for i <- 1..10 do
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A#{i}", "#{i * 10}")
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B#{i}", "Value #{i}")
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C#{i}", "#{i * 5}")
    end

    # Add a cell value rule (equal to)
    ConditionalFormatting.add_cell_value_rule(
      spreadsheet,
      "Sheet1",
      "A1:A10",
      "equal",
      "50",
      nil,
      # Red
      "#FF0000"
    )

    # Add a color scale
    ConditionalFormatting.add_color_scale(
      spreadsheet,
      "Sheet1",
      "A1:A10",
      nil,
      nil,
      # Blue for min
      "#0000FF",
      nil,
      nil,
      # Green for max
      "#00FF00"
    )

    # Add a data bar
    ConditionalFormatting.add_data_bar(
      spreadsheet,
      "Sheet1",
      "C1:C10",
      nil,
      nil,
      # Blue data bars
      "#638EC6"
    )

    # Add top 3 rule
    ConditionalFormatting.add_top_bottom_rule(
      spreadsheet,
      "Sheet1",
      "A1:A10",
      "top",
      3,
      false,
      # Yellow
      "#FFFF00"
    )

    # Add text rule begins with
    ConditionalFormatting.add_text_rule(
      spreadsheet,
      "Sheet1",
      "B1:B10",
      "beginsWith",
      "Value 1",
      # Orange
      "#FFC000"
    )

    # Add above average rule
    ConditionalFormatting.add_above_below_average_rule(
      spreadsheet,
      "Sheet1",
      "C1:C10",
      "above",
      nil,
      # Green
      "#00FF00"
    )

    # Save the spreadsheet
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)
    assert File.exists?(@output_path)

    # Now test getting all conditional formatting rules
    all_rules = ConditionalFormatting.get_conditional_formatting_rules(spreadsheet, "Sheet1")
    assert is_list(all_rules)
    # We added 6 rules
    assert length(all_rules) == 6

    # Test getting only cell value rules
    cell_value_rules = ConditionalFormatting.get_cell_value_rules(spreadsheet, "Sheet1")
    assert is_list(cell_value_rules)
    assert length(cell_value_rules) == 1

    # Verify cell value rule properties
    [rule] = cell_value_rules
    assert rule.range == "A1:A10"
    assert rule.formula == "50"

    # Test getting color scales
    {:ok, color_scales} = ConditionalFormatting.get_color_scales(spreadsheet, "Sheet1")
    assert is_list(color_scales)
    assert length(color_scales) == 1

    # Test getting data bars
    {:ok, data_bars} = ConditionalFormatting.get_data_bars(spreadsheet, "Sheet1")
    assert is_list(data_bars)
    assert length(data_bars) == 1
    assert hd(data_bars).range == "C1:C10"

    # Test getting top/bottom rules
    top_bottom_rules = ConditionalFormatting.get_top_bottom_rules(spreadsheet, "Sheet1")
    assert is_list(top_bottom_rules)
    assert length(top_bottom_rules) == 1
    assert hd(top_bottom_rules).range == "A1:A10"
    assert hd(top_bottom_rules).rank == 3

    # Test getting text rules
    text_rules = ConditionalFormatting.get_text_rules(spreadsheet, "Sheet1")
    assert is_list(text_rules)
    assert length(text_rules) == 1
    assert hd(text_rules).range == "B1:B10"
    assert hd(text_rules).text == "Value 1"

    # Test getting above/below average rules
    avg_rules = ConditionalFormatting.get_above_below_average_rules(spreadsheet, "Sheet1")
    assert is_list(avg_rules)
    assert length(avg_rules) == 1
    assert hd(avg_rules).range == "C1:C10"
  end

  test "get conditional formatting rules with filtering by range" do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add test data
    for i <- 1..10 do
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A#{i}", "#{i * 10}")
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B#{i}", "#{i * 5}")
    end

    # Add rules to different ranges
    ConditionalFormatting.add_cell_value_rule(
      spreadsheet,
      "Sheet1",
      "A1:A5",
      "greaterThan",
      "30",
      nil,
      # Red
      "#FF0000"
    )

    ConditionalFormatting.add_cell_value_rule(
      spreadsheet,
      "Sheet1",
      "A6:A10",
      "greaterThan",
      "70",
      nil,
      # Blue
      "#0000FF"
    )

    ConditionalFormatting.add_data_bar(
      spreadsheet,
      "Sheet1",
      "B1:B10",
      nil,
      nil,
      # Green
      "#00FF00"
    )

    # Filter rules by range
    a1_a5_rules =
      ConditionalFormatting.get_conditional_formatting_rules(spreadsheet, "Sheet1", "A1:A5")

    assert is_list(a1_a5_rules)
    assert length(a1_a5_rules) == 1
    assert hd(a1_a5_rules).range == "A1:A5"

    a6_a10_rules =
      ConditionalFormatting.get_conditional_formatting_rules(spreadsheet, "Sheet1", "A6:A10")

    assert is_list(a6_a10_rules)
    assert length(a6_a10_rules) == 1
    assert hd(a6_a10_rules).range == "A6:A10"

    # Test filter by rule type and range
    a1_a5_cell_rules = ConditionalFormatting.get_cell_value_rules(spreadsheet, "Sheet1", "A1:A5")
    assert length(a1_a5_cell_rules) == 1
    assert hd(a1_a5_cell_rules).formula == "30"
  end

  test "read from and write to existing file with conditional formatting" do
    # First create a file with conditional formatting
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add test data
    for i <- 1..5 do
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A#{i}", "#{i * 10}")
    end

    # Add a data bar rule
    ConditionalFormatting.add_data_bar(
      spreadsheet,
      "Sheet1",
      "A1:A5",
      nil,
      nil,
      # Red
      "#FF0000"
    )

    # Save the file
    test_file = "#{@output_path}.existing.xlsx"
    assert :ok = UmyaSpreadsheet.write(spreadsheet, test_file)

    # Now read the file back and check formatting rules
    {:ok, read_spreadsheet} = UmyaSpreadsheet.read(test_file)
    rules = ConditionalFormatting.get_conditional_formatting_rules(read_spreadsheet, "Sheet1")

    assert length(rules) == 1
    assert hd(rules).range == "A1:A5"

    {:ok, data_bars} = ConditionalFormatting.get_data_bars(read_spreadsheet, "Sheet1")
    assert length(data_bars) == 1
    assert hd(data_bars).range == "A1:A5"

    # Add another rule and save again
    ConditionalFormatting.add_color_scale(
      read_spreadsheet,
      "Sheet1",
      "A1:A5",
      nil,
      nil,
      # Blue for min
      "#0000FF",
      nil,
      nil,
      # Green for max
      "#00FF00"
    )

    assert :ok = UmyaSpreadsheet.write(read_spreadsheet, "#{test_file}.updated.xlsx")

    # Read the updated file and verify both rules exist
    {:ok, updated_spreadsheet} = UmyaSpreadsheet.read("#{test_file}.updated.xlsx")

    updated_rules =
      ConditionalFormatting.get_conditional_formatting_rules(updated_spreadsheet, "Sheet1")

    assert length(updated_rules) == 2
  end
end
