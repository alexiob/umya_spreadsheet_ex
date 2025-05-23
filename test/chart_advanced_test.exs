defmodule UmyaSpreadsheet.ChartAdvancedTest do
  use ExUnit.Case, async: true

  test "3D chart with advanced options" do
    # Create new spreadsheet
    {:ok, sheet} = UmyaSpreadsheet.new()

    # Set some data in a sheet for the chart
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "A1", "Quarter")
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "B1", "Sales")
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "C1", "Profit")
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "D1", "Expenses")

    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "A2", "Q1")
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "B2", 10000)
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "C2", 5000)
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "D2", 5000)

    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "A3", "Q2")
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "B3", 12000)
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "C3", 6000)
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "D3", 6000)

    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "A4", "Q3")
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "B4", 15000)
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "C4", 7500)
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "D4", 7500)

    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "A5", "Q4")
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "B5", 18000)
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "C5", 9000)
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "D5", 9000)

    # Create data series for the chart
    data_series = ["Sheet1!$B$2:$B$5", "Sheet1!$C$2:$C$5", "Sheet1!$D$2:$D$5"]
    series_titles = ["Sales", "Profit", "Expenses"]
    point_titles = ["Q1", "Q2", "Q3", "Q4"]

    # Test all chart types with advanced options

    # 1. 3D Bar Chart
    :ok =
      UmyaSpreadsheet.add_chart_with_options(
        sheet,
        "Sheet1",
        "Bar3DChart",
        "F2",
        "L12",
        "Quarterly Financial Results (3D Bar)",
        data_series,
        series_titles,
        point_titles,
        # style
        10,
        # vary_colors
        true,
        # 3D view options
        %{rot_x: 20, rot_y: 30, perspective: 30, height_percent: 150},
        # legend options
        %{position: :right, overlay: false},
        # axes options (not implemented)
        nil,
        # data labels
        %{
          show_values: true,
          show_percent: false,
          show_category_name: false,
          show_series_name: false
        }
      )

    # 2. 3D Pie Chart
    :ok =
      UmyaSpreadsheet.add_chart_with_options(
        sheet,
        "Sheet1",
        "Pie3DChart",
        "F13",
        "L23",
        "Q4 Financial Breakdown (3D Pie)",
        # Just using Q4 data for the pie
        ["Sheet1!$B$5:$D$5"],
        ["Q4 Financials"],
        ["Sales", "Profit", "Expenses"],
        # style
        12,
        # vary_colors
        true,
        # 3D view options
        %{rot_x: 40, rot_y: 30, perspective: 30, height_percent: 150},
        # legend options
        %{position: :bottom, overlay: false},
        # axes options (not implemented)
        nil,
        # data labels
        %{
          show_values: true,
          show_percent: true,
          show_category_name: true,
          show_series_name: false
        }
      )

    # 3. Line Chart
    :ok =
      UmyaSpreadsheet.add_chart_with_options(
        sheet,
        "Sheet1",
        "LineChart",
        "M2",
        "S12",
        "Quarterly Trends",
        data_series,
        series_titles,
        point_titles,
        # style
        15,
        # vary_colors
        false,
        # 3D view options (not applicable)
        nil,
        # legend options
        %{position: :top_right, overlay: true},
        # axes options (not implemented)
        nil,
        # data labels
        %{
          show_values: false,
          show_percent: false,
          show_category_name: false,
          show_series_name: false
        }
      )

    # 4. 3D Line Chart
    :ok =
      UmyaSpreadsheet.add_chart_with_options(
        sheet,
        "Sheet1",
        "Line3DChart",
        "M13",
        "S23",
        "Quarterly Trends (3D)",
        data_series,
        series_titles,
        point_titles,
        # style
        20,
        # vary_colors
        false,
        # 3D view options
        %{rot_x: 15, rot_y: 20, perspective: 30, height_percent: 120},
        # legend options
        %{position: :top, overlay: false},
        # axes options (not implemented)
        nil,
        # data labels
        %{
          show_values: true,
          show_percent: false,
          show_category_name: false,
          show_series_name: false
        }
      )

    # Save file for manual inspection
    test_output_path = "test/result_files/chart_advanced_test_output.xlsx"
    :ok = UmyaSpreadsheet.write(sheet, test_output_path)

    # Verify file creation
    assert File.exists?(test_output_path)
  end
end
