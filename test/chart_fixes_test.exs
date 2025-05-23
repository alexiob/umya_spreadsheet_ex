defmodule UmyaSpreadsheet.ChartFixesTest do
  use ExUnit.Case, async: true

  test "chart functionality" do
    # Create new spreadsheet
    {:ok, sheet} = UmyaSpreadsheet.new()

    # Set some data in a sheet for the chart
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "A1", "Months")
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "B1", "Sales")
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "C1", "Profit")

    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "A2", "January")
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "B2", 5000)
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "C2", 1000)

    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "A3", "February")
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "B3", 7000)
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "C3", 1500)

    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "A4", "March")
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "B4", 8000)
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "C4", 2000)

    # Test basic chart creation
    data_series = ["Sheet1!$B$2:$B$4", "Sheet1!$C$2:$C$4"]
    series_titles = ["Sales", "Profit"]

    # Create a line chart
    :ok =
      UmyaSpreadsheet.add_chart(
        sheet,
        "Sheet1",
        "LineChart",
        "E2",
        "K15",
        "Sales and Profit by Month",
        data_series,
        series_titles,
        []
      )

    # Change chart legend position
    :ok =
      UmyaSpreadsheet.set_chart_legend_position(
        sheet,
        "Sheet1",
        # First chart (index 0)
        0,
        "bottom",
        # Not overlaid
        false
      )

    # Set chart axis titles
    :ok =
      UmyaSpreadsheet.set_chart_axis_titles(
        sheet,
        "Sheet1",
        0,
        "Month",
        "Amount ($)"
      )

    # Set data labels
    :ok =
      UmyaSpreadsheet.set_chart_data_labels(
        sheet,
        "Sheet1",
        0,
        # Show values
        true,
        # Don't show percentages
        false,
        # Don't show category names
        false,
        # Don't show series names
        false,
        "center"
      )

    # Create a second chart - Bar Chart
    :ok =
      UmyaSpreadsheet.add_chart(
        sheet,
        "Sheet1",
        "BarChart",
        "E17",
        "K30",
        "Bar Chart",
        data_series,
        series_titles,
        []
      )

    # Create a third chart - Pie Chart
    :ok =
      UmyaSpreadsheet.add_chart(
        sheet,
        "Sheet1",
        "PieChart",
        "M2",
        "S15",
        "Sales Distribution",
        ["Sheet1!$B$2:$B$4"],
        ["Sales"],
        ["January", "February", "March"]
      )

    # Save the file
    test_output_path = "test/result_files/test_chart_output.xlsx"
    :ok = UmyaSpreadsheet.write(sheet, test_output_path)

    # Verify file was created
    assert File.exists?(test_output_path)
  end

  # Test advanced chart creation with options
  test "chart advanced options" do
    {:ok, sheet} = UmyaSpreadsheet.new()

    # Set some data in a sheet for the chart
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "A1", "Months")
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "B1", "Sales")

    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "A2", "January")
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "B2", 5000)

    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "A3", "February")
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "B3", 7000)

    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "A4", "March")
    UmyaSpreadsheet.set_cell_value(sheet, "Sheet1", "B4", 8000)

    data_series = ["Sheet1!B2:B4"]
    series_titles = ["Sales"]

    # Create a 3D line chart with advanced options
    :ok =
      UmyaSpreadsheet.add_chart_with_options(
        sheet,
        "Sheet1",
        "Line3DChart",
        "E2",
        "K15",
        "Sales by Month",
        data_series,
        series_titles,
        # point_titles
        [],
        # style
        15,
        # vary_colors
        true,
        # 3D view options
        %{rot_x: 30, rot_y: 20, perspective: 30, height_percent: 150},
        # legend options
        %{position: :bottom, overlay: false},
        # axes options (not implemented)
        nil,
        # data labels
        %{show_values: true, show_percent: false},
        # chart-specific options (not implemented)
        nil
      )

    # Save the advanced charts file
    test_output_path = "test/result_files/test_chart_advanced_output.xlsx"
    :ok = UmyaSpreadsheet.write(sheet, test_output_path)
  end
end
