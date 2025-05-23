defmodule UmyaSpreadsheet.ChartTest do
  use ExUnit.Case, async: true

  @output_path "test/result_files/charts_output.xlsx"

  setup do
    # Create a new spreadsheet for testing
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Prepare data for charts
    # Header row
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Category")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Series 1")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C1", "Series 2")

    # Data rows
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "Q1")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", "10")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C2", "8")

    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "Q2")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B3", "12")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C3", "6")

    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A4", "Q3")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B4", "8")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C4", "10")

    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A5", "Q4")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B5", "15")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C5", "12")

    # Clean up any previous test result files
    File.rm(@output_path)

    %{spreadsheet: spreadsheet}
  end

  test "add line chart", %{spreadsheet: spreadsheet} do
    # Define chart data
    data_series = ["Sheet1!$B$2:$B$5", "Sheet1!$C$2:$C$5"]
    series_titles = ["Revenue", "Expenses"]
    point_titles = ["Q1", "Q2", "Q3", "Q4"]

    # Add a line chart
    assert :ok =
             UmyaSpreadsheet.add_chart(
               spreadsheet,
               "Sheet1",
               "LineChart",
               "E1",
               "J10",
               "Quarterly Performance",
               data_series,
               series_titles,
               point_titles
             )

    # Save the spreadsheet
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)
  end

  test "add pie chart", %{spreadsheet: spreadsheet} do
    # Define chart data
    data_series = ["Sheet1!$B$2:$B$5"]
    series_titles = ["Revenue"]
    point_titles = ["Q1", "Q2", "Q3", "Q4"]

    # Add a pie chart
    assert :ok =
             UmyaSpreadsheet.add_chart(
               spreadsheet,
               "Sheet1",
               "PieChart",
               "E12",
               "J20",
               "Revenue Distribution",
               data_series,
               series_titles,
               point_titles
             )

    # Save the spreadsheet
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)
  end

  test "add bar chart", %{spreadsheet: spreadsheet} do
    # Define chart data
    data_series = ["Sheet1!$B$2:$B$5", "Sheet1!$C$2:$C$5"]
    series_titles = ["Revenue", "Expenses"]
    point_titles = ["Q1", "Q2", "Q3", "Q4"]

    # Add a bar chart
    assert :ok =
             UmyaSpreadsheet.add_chart(
               spreadsheet,
               "Sheet1",
               "BarChart",
               "K1",
               "P10",
               "Quarterly Comparison",
               data_series,
               series_titles,
               point_titles
             )

    # Save the spreadsheet
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)
  end
end
