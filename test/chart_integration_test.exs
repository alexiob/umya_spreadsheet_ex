defmodule UmyaSpreadsheet.ChartIntegrationTest do
  use ExUnit.Case, async: true

  @output_path "test/result_files/chart_integration_test.xlsx"

  setup do
    # Create a new spreadsheet for testing
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add data sheet for chart data
    :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "DataSheet")

    # Add data for charts
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "DataSheet", "G7", "100")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "DataSheet", "G8", "200")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "DataSheet", "G9", "300")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "DataSheet", "G10", "400")

    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "DataSheet", "H7", "150")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "DataSheet", "H8", "250")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "DataSheet", "H9", "350")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "DataSheet", "H10", "450")

    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "DataSheet", "I7", "120")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "DataSheet", "I8", "220")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "DataSheet", "I9", "320")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "DataSheet", "I10", "420")

    # Add a chart sheet
    :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "ChartSheet")

    # Clean up any previous test result files
    File.rm(@output_path)

    %{spreadsheet: spreadsheet}
  end

  test "add different chart types to a spreadsheet", %{spreadsheet: spreadsheet} do
    # Add line chart
    data_series = ["DataSheet!$G$7:$G$10", "DataSheet!$H$7:$H$10"]
    series_titles = ["Line1", "Line2"]
    point_titles = ["Point1", "Point2", "Point3", "Point4"]

    assert :ok =
             UmyaSpreadsheet.add_chart(
               spreadsheet,
               "ChartSheet",
               "LineChart",
               "A1",
               "B2",
               "Line Chart Title",
               data_series,
               series_titles,
               point_titles
             )

    # Add pie chart
    assert :ok =
             UmyaSpreadsheet.add_chart(
               spreadsheet,
               "ChartSheet",
               "PieChart",
               "C1",
               "D2",
               "Pie Chart Title",
               data_series,
               series_titles,
               point_titles
             )

    # Add bar chart
    assert :ok =
             UmyaSpreadsheet.add_chart(
               spreadsheet,
               "ChartSheet",
               "BarChart",
               "E1",
               "F2",
               "Bar Chart Title",
               data_series,
               series_titles,
               point_titles
             )

    # Add doughnut chart
    assert :ok =
             UmyaSpreadsheet.add_chart(
               spreadsheet,
               "ChartSheet",
               "DoughnutChart",
               "A3",
               "B4",
               "Doughnut Chart Title",
               data_series,
               series_titles,
               point_titles
             )

    # Add area chart
    assert :ok =
             UmyaSpreadsheet.add_chart(
               spreadsheet,
               "ChartSheet",
               "AreaChart",
               "C3",
               "D4",
               "Area Chart Title",
               data_series,
               series_titles,
               point_titles
             )

    # Add 3D charts
    assert :ok =
             UmyaSpreadsheet.add_chart(
               spreadsheet,
               "ChartSheet",
               "Bar3DChart",
               "E3",
               "F4",
               "3D Bar Chart Title",
               data_series,
               series_titles,
               point_titles
             )

    # Add radar chart
    data_series_3 = ["DataSheet!$G$7:$G$10", "DataSheet!$H$7:$H$10", "DataSheet!$I$7:$I$10"]
    series_titles_3 = ["Line1", "Line2", "Line3"]

    assert :ok =
             UmyaSpreadsheet.add_chart(
               spreadsheet,
               "ChartSheet",
               "RadarChart",
               "A5",
               "B6",
               "Radar Chart Title",
               data_series_3,
               series_titles_3,
               point_titles
             )

    # Write the spreadsheet to file
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)

    # Verify the file was created
    assert File.exists?(@output_path)

    # Read back the file to verify it can be opened
    {:ok, _read_spreadsheet} = UmyaSpreadsheet.read(@output_path)
  end
end
