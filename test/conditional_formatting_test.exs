defmodule UmyaSpreadsheet.ConditionalFormattingTest do
  use ExUnit.Case, async: true

  alias UmyaSpreadsheet

  @output_path "test/result_files/conditional_formatting_test.xlsx"

  setup_all do
    # Make sure the result files directory exists
    File.mkdir_p!("test/result_files")
    :ok
  end

  test "add_cell_value_rule with equal operator" do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add test data
    for i <- 1..10 do
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A#{i}", "#{i * 10}")
    end

    # Add conditional formatting for cells equal to 50
    assert :ok =
             UmyaSpreadsheet.add_cell_value_rule(
               spreadsheet,
               "Sheet1",
               "A1:A10",
               "equal",
               "50",
               nil,
               # Red background
               "#FF0000"
             )

    # Save the spreadsheet
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)
    assert File.exists?(@output_path)
  end

  test "add_cell_value_rule with between operator" do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add test data
    for i <- 1..10 do
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A#{i}", "#{i * 10}")
    end

    # Add conditional formatting for cells between 30 and 70
    assert :ok =
             UmyaSpreadsheet.add_cell_value_rule(
               spreadsheet,
               "Sheet1",
               "A1:A10",
               "between",
               "30",
               "70",
               # Green background
               "#00FF00"
             )

    # Save the spreadsheet
    assert :ok = UmyaSpreadsheet.write(spreadsheet, "#{@output_path}.between.xlsx")
    assert File.exists?("#{@output_path}.between.xlsx")
  end

  test "add_color_scale with two colors" do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add test data
    for i <- 1..10 do
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A#{i}", "#{i * 10}")
    end

    # Add two-color scale (min to max)
    assert :ok =
             UmyaSpreadsheet.add_color_scale(
               spreadsheet,
               "Sheet1",
               "A1:A10",
               # Use default min
               nil,
               # No min value
               nil,
               # No min color, use default white
               nil,
               # Use default max
               nil,
               # No max value
               nil,
               # Use a struct for color
               %{argb: "FF00FF00"}  # Green for maximum
             )

    # Save the spreadsheet
    assert :ok = UmyaSpreadsheet.write(spreadsheet, "#{@output_path}.2color.xlsx")
    assert File.exists?("#{@output_path}.2color.xlsx")
  end

  test "add_color_scale with three colors" do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add test data
    for i <- 1..10 do
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A#{i}", "#{i * 10}")
    end

    # Add three-color scale with custom midpoint
    assert :ok =
             UmyaSpreadsheet.add_color_scale(
               spreadsheet,
               "Sheet1",
               "A1:A10",
               # Use minimum value
               "min",
               # No min value
               nil,
               # White for minimum 
               %{argb: "FFFFFFFF"},
               # Use percentile for mid-point
               "percentile",
               # 50th percentile
               "50",
               # Yellow for middle
               %{argb: "FFFFFF00"},
               # Use maximum value
               "max",
               # No max value
               nil,
               # Green for maximum
               %{argb: "FF00FF00"}
             )

    # Save the spreadsheet
    assert :ok = UmyaSpreadsheet.write(spreadsheet, "#{@output_path}.3color.xlsx")
    assert File.exists?("#{@output_path}.3color.xlsx")
  end

  test "add_data_bar" do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add test data
    for i <- 1..10 do
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A#{i}", "#{i * 10}")
    end

    # Add data bars
    assert :ok =
             UmyaSpreadsheet.add_data_bar(
               spreadsheet,
               "Sheet1",
               "A1:A10",
               # Use default min
               nil,
               # Use default max
               nil,
               # Blue bars
               "#6C8EBF"
             )

    # Save the spreadsheet
    assert :ok = UmyaSpreadsheet.write(spreadsheet, "#{@output_path}.databars.xlsx")
    assert File.exists?("#{@output_path}.databars.xlsx")
  end

  test "add_top_bottom_rule for top values" do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add test data with some random values
    values = [85, 42, 91, 38, 67, 83, 75, 28, 50, 96]

    for {value, i} <- Enum.with_index(values, 1) do
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A#{i}", "#{value}")
    end

    # Highlight top 3 values
    assert :ok =
             UmyaSpreadsheet.add_top_bottom_rule(
               spreadsheet,
               "Sheet1",
               "A1:A10",
               "top",
               # Top 3 values
               3,
               # Not percentage
               false,
               # Red highlight
               "#FF0000"
             )

    # Save the spreadsheet
    assert :ok = UmyaSpreadsheet.write(spreadsheet, "#{@output_path}.top3.xlsx")
    assert File.exists?("#{@output_path}.top3.xlsx")
  end

  test "add_top_bottom_rule for bottom percentage" do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add test data with some random values
    values = [85, 42, 91, 38, 67, 83, 75, 28, 50, 96]

    for {value, i} <- Enum.with_index(values, 1) do
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A#{i}", "#{value}")
    end

    # Highlight bottom 30%
    assert :ok =
             UmyaSpreadsheet.add_top_bottom_rule(
               spreadsheet,
               "Sheet1",
               "A1:A10",
               "bottom",
               # Bottom 30%
               30,
               # Use as percentage
               true,
               # Yellow highlight
               "#FFFF00"
             )

    # Save the spreadsheet
    assert :ok = UmyaSpreadsheet.write(spreadsheet, "#{@output_path}.bottom30pct.xlsx")
    assert File.exists?("#{@output_path}.bottom30pct.xlsx")
  end

  test "add_text_rule with contains" do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add test data
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Error in file")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "Warning message")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "Success notification")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A4", "Error: file not found")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A5", "Processing complete")

    # Highlight cells containing "Error"
    assert :ok =
             UmyaSpreadsheet.add_text_rule(
               spreadsheet,
               "Sheet1",
               "A1:A10",
               "contains",
               "Error",
               # Red highlight
               "#FF0000"
             )

    # Save the spreadsheet
    assert :ok = UmyaSpreadsheet.write(spreadsheet, "#{@output_path}.text.xlsx")
    assert File.exists?("#{@output_path}.text.xlsx")
  end

  test "comprehensive test with multiple conditional formatting rules" do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Sheet1: Cell value rules
    UmyaSpreadsheet.add_sheet(spreadsheet, "Values")

    for i <- 1..20 do
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Values", "A#{i}", "#{i * 5}")
    end

    # Less than 30 (red)
    UmyaSpreadsheet.add_cell_value_rule(
      spreadsheet,
      "Values",
      "A1:A20",
      "lessThan",
      "30",
      nil,
      "#FF0000"
    )

    # Between 30 and 60 (yellow)
    UmyaSpreadsheet.add_cell_value_rule(
      spreadsheet,
      "Values",
      "A1:A20",
      "between",
      "30",
      "60",
      "#FFFF00"
    )

    # Greater than 60 (green)
    UmyaSpreadsheet.add_cell_value_rule(
      spreadsheet,
      "Values",
      "A1:A20",
      "greaterThan",
      "60",
      nil,
      "#00FF00"
    )

    # Sheet2: Color scales
    UmyaSpreadsheet.add_sheet(spreadsheet, "ColorScales")

    for i <- 1..20 do
      UmyaSpreadsheet.set_cell_value(spreadsheet, "ColorScales", "A#{i}", "#{i * 5}")
    end

    # Three color scale
    UmyaSpreadsheet.add_color_scale(
      spreadsheet,
      "ColorScales",
      "A1:A20",
      "min",
      nil,
      %{argb: "FFFF0000"}, # Red
      "percentile",
      "50",
      %{argb: "FFFFFF00"}, # Yellow
      "max",
      nil,
      %{argb: "FF00FF00"} # Green
    )

    # Sheet3: Data bars
    UmyaSpreadsheet.add_sheet(spreadsheet, "DataBars")

    for i <- 1..20 do
      UmyaSpreadsheet.set_cell_value(spreadsheet, "DataBars", "A#{i}", "#{i * 5}")
    end

    # Data bars
    UmyaSpreadsheet.add_data_bar(
      spreadsheet,
      "DataBars",
      "A1:A20",
      nil,
      nil,
      "#6C8EBF"
    )

    # Sheet4: Text rules
    UmyaSpreadsheet.add_sheet(spreadsheet, "TextRules")

    statuses = [
      "Error: File not found",
      "Warning: Disk space low",
      "Success: Upload complete",
      "Info: Processing started",
      "Error: Connection failed",
      "Warning: Timeout approaching",
      "Success: Task completed",
      "Info: 20 items processed",
      "Error: Invalid format",
      "Warning: Battery low"
    ]

    for {status, i} <- Enum.with_index(statuses, 1) do
      UmyaSpreadsheet.set_cell_value(spreadsheet, "TextRules", "A#{i}", status)
    end

    # Contains "Error" (red)
    UmyaSpreadsheet.add_text_rule(
      spreadsheet,
      "TextRules",
      "A1:A10",
      "contains",
      "Error",
      "#FF0000"
    )

    # Contains "Warning" (yellow)
    UmyaSpreadsheet.add_text_rule(
      spreadsheet,
      "TextRules",
      "A1:A10",
      "contains",
      "Warning",
      "#FFFF00"
    )

    # Contains "Success" (green)
    UmyaSpreadsheet.add_text_rule(
      spreadsheet,
      "TextRules",
      "A1:A10",
      "contains",
      "Success",
      "#00FF00"
    )

    # Sheet5: Top/Bottom rules
    UmyaSpreadsheet.add_sheet(spreadsheet, "TopBottom")
    values = [85, 42, 91, 38, 67, 83, 75, 28, 50, 96, 62, 39, 88, 71, 44, 53, 79, 33, 59, 82]

    for {value, i} <- Enum.with_index(values, 1) do
      UmyaSpreadsheet.set_cell_value(spreadsheet, "TopBottom", "A#{i}", "#{value}")
    end

    # Top 3 values (green)
    UmyaSpreadsheet.add_top_bottom_rule(
      spreadsheet,
      "TopBottom",
      "A1:A20",
      "top",
      3,
      false,
      "#00FF00"
    )

    # Bottom 20% (red)
    UmyaSpreadsheet.add_top_bottom_rule(
      spreadsheet,
      "TopBottom",
      "A1:A20",
      "bottom",
      20,
      true,
      "#FF0000"
    )

    # Save the spreadsheet
    assert :ok = UmyaSpreadsheet.write(spreadsheet, "#{@output_path}.comprehensive.xlsx")
    assert File.exists?("#{@output_path}.comprehensive.xlsx")
  end
end
