defmodule UmyaSpreadsheet.ConditionalFormattingDirectTest do
  use ExUnit.Case, async: true

  @moduledoc """
  This module tests the conditional formatting functions directly
  with working examples to help debug issues.
  """

  alias UmyaSpreadsheetEx.CustomStructs.CustomColor

  @output_path "test/result_files/conditional_format_direct_test"

  test "direct test for add_color_scale with explicit format" do
    # Create a new spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Test data
    sheet_name = "Sheet1"
    for i <- 1..10 do
      UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, "A#{i}", "#{i * 10}")
    end

    # Define color with proper format
    min_color = %CustomColor{argb: "FFFFFFFF"}  # White
    max_color = %CustomColor{argb: "FF00FF00"}  # Green

    # Call directly
    ref = spreadsheet.reference

    # Get the function
    {:ok, true} = UmyaNative.add_color_scale(
      ref,
      sheet_name,
      "A1:A10",
      "min",
      nil,
      min_color,
      nil,
      nil,
      nil,
      "max",
      nil,
      max_color
    )

    # Save and check result
    assert :ok = UmyaSpreadsheet.write(spreadsheet, "#{@output_path}.direct_test.xlsx")
    assert File.exists?("#{@output_path}.direct_test.xlsx")
  end
end
