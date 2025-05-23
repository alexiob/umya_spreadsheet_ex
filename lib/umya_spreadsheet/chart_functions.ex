defmodule UmyaSpreadsheet.ChartFunctions do
  @moduledoc """
  Functions for creating and manipulating charts in a spreadsheet.
  """
  
  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaNative

  @doc """
  Adds a chart to a spreadsheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `chart_type` - The type of chart (e.g., "bar", "line", "pie")
  - `from_cell` - The top-left cell of the chart position
  - `to_cell` - The bottom-right cell of the chart position
  - `title` - The title of the chart
  - `data_series` - A list of data series ranges
  - `series_titles` - A list of series titles
  - `point_titles` - A list of point titles (categories)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.ChartFunctions.add_chart(
        spreadsheet,
        "Sheet1",
        "bar",
        "E2",
        "K15",
        "Sales Chart",
        ["B2:B10"],
        ["Sales"],
        ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep"]
      )
  """
  def add_chart(
        %Spreadsheet{reference: ref},
        sheet_name,
        chart_type,
        from_cell,
        to_cell,
        title,
        data_series,
        series_titles,
        point_titles
      ) do
    case UmyaNative.add_chart(
           ref,
           sheet_name,
           chart_type,
           from_cell,
           to_cell,
           title,
           data_series,
           series_titles,
           point_titles
         ) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Adds a chart to a spreadsheet with advanced options.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `chart_type` - The type of chart (e.g., "bar", "line", "pie")
  - `from_cell` - The top-left cell of the chart position
  - `to_cell` - The bottom-right cell of the chart position
  - `title` - The title of the chart
  - `data_series` - A list of data series ranges
  - `series_titles` - A list of series titles
  - `point_titles` - A list of point titles (categories)
  - `style` - Chart style number
  - `vary_colors` - Boolean indicating whether to vary colors by point
  - `view_3d` - 3D view options (rotation, perspective)
  - `legend` - Legend options (position, overlay)
  - `axes` - Axes options (titles, formatting)
  - `data_labels` - Data label options

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.ChartFunctions.add_chart_with_options(
        spreadsheet,
        "Sheet1",
        "bar",
        "E2",
        "K15",
        "Sales Chart",
        ["B2:B10"],
        ["Sales"],
        ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep"],
        2,
        true,
        %{rot_x: 30, rot_y: 20, perspective: 30, height_percent: 100},
        %{position: "right", overlay: false},
        %{category_axis_title: "Month", value_axis_title: "Amount"},
        %{show_values: true, show_percent: false, show_category_name: false, show_series_name: true, position: "outside_end"}
      )
  """
  def add_chart_with_options(
        %Spreadsheet{} = spreadsheet,
        sheet_name,
        chart_type,
        from_cell,
        to_cell,
        title,
        data_series,
        series_titles,
        point_titles,
        style,
        vary_colors,
        view_3d,
        legend,
        axes,
        data_labels
      ) do
    add_chart_with_options(
      spreadsheet,
      sheet_name,
      chart_type,
      from_cell,
      to_cell,
      title,
      data_series,
      series_titles,
      point_titles,
      style,
      vary_colors,
      view_3d,
      legend,
      axes,
      data_labels,
      nil
    )
  end

  @doc """
  Adds a chart to a spreadsheet with advanced options, including chart-specific options.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `chart_type` - The type of chart (e.g., "bar", "line", "pie")
  - `from_cell` - The top-left cell of the chart position
  - `to_cell` - The bottom-right cell of the chart position
  - `title` - The title of the chart
  - `data_series` - A list of data series ranges
  - `series_titles` - A list of series titles
  - `point_titles` - A list of point titles (categories)
  - `style` - Chart style number
  - `vary_colors` - Boolean indicating whether to vary colors by point
  - `view_3d` - 3D view options (rotation, perspective)
  - `legend` - Legend options (position, overlay)
  - `axes` - Axes options (titles, formatting)
  - `data_labels` - Data label options
  - `chart_specific` - Chart-specific options

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure
  """
  def add_chart_with_options(
        %Spreadsheet{reference: ref},
        sheet_name,
        chart_type,
        from_cell,
        to_cell,
        title,
        data_series,
        series_titles,
        point_titles,
        style,
        vary_colors,
        view_3d,
        legend,
        axes,
        data_labels,
        chart_specific
      ) do
    case UmyaNative.add_chart_with_options(
           ref,
           sheet_name,
           chart_type,
           from_cell,
           to_cell,
           title,
           data_series,
           series_titles,
           point_titles,
           style,
           vary_colors,
           view_3d,
           legend,
           axes,
           data_labels,
           chart_specific
         ) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets the chart style.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `chart_index` - The index of the chart (0-based)
  - `style` - The style number (1-48)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.ChartFunctions.set_chart_style(spreadsheet, "Sheet1", 0, 5)
  """
  def set_chart_style(%Spreadsheet{reference: ref}, sheet_name, chart_index, style) do
    case UmyaNative.set_chart_style(ref, sheet_name, chart_index, style) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets data labels for a chart.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `chart_index` - The index of the chart (0-based)
  - `show_values` - Boolean indicating whether to show values
  - `show_percent` - Boolean indicating whether to show percentages
  - `show_category_name` - Boolean indicating whether to show category names
  - `show_series_name` - Boolean indicating whether to show series names
  - `position` - The position of the data labels (e.g., "outside_end", "center", "inside_end")

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.ChartFunctions.set_chart_data_labels(
        spreadsheet, 
        "Sheet1", 
        0, 
        true, 
        false, 
        true, 
        false, 
        "outside_end"
      )
  """
  def set_chart_data_labels(
        %Spreadsheet{reference: ref},
        sheet_name,
        chart_index,
        show_values,
        show_percent,
        show_category_name,
        show_series_name,
        position
      ) do
    case UmyaNative.set_chart_data_labels(
           ref,
           sheet_name,
           chart_index,
           show_values,
           show_percent,
           show_category_name,
           show_series_name,
           position
         ) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets the legend position for a chart.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `chart_index` - The index of the chart (0-based)
  - `position` - The position of the legend (e.g., "right", "left", "top", "bottom", "top_right")
  - `overlay` - Boolean indicating whether the legend should overlay the chart

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.ChartFunctions.set_chart_legend_position(spreadsheet, "Sheet1", 0, "right", false)
  """
  def set_chart_legend_position(
        %Spreadsheet{reference: ref},
        sheet_name,
        chart_index,
        position,
        overlay
      ) do
    case UmyaNative.set_chart_legend_position(ref, sheet_name, chart_index, position, overlay) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets 3D view properties for a chart.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `chart_index` - The index of the chart (0-based)
  - `rot_x` - X-axis rotation (degrees)
  - `rot_y` - Y-axis rotation (degrees)
  - `perspective` - Perspective (0-100)
  - `height_percent` - Height as percentage of width

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.ChartFunctions.set_chart_3d_view(spreadsheet, "Sheet1", 0, 30, 20, 30, 100)
  """
  def set_chart_3d_view(
        %Spreadsheet{reference: ref},
        sheet_name,
        chart_index,
        rot_x,
        rot_y,
        perspective,
        height_percent
      ) do
    case UmyaNative.set_chart_3d_view(
           ref,
           sheet_name,
           chart_index,
           rot_x,
           rot_y,
           perspective,
           height_percent
         ) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Sets axis titles for a chart.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `chart_index` - The index of the chart (0-based)
  - `category_axis_title` - The title for the category (X) axis
  - `value_axis_title` - The title for the value (Y) axis

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.ChartFunctions.set_chart_axis_titles(spreadsheet, "Sheet1", 0, "Month", "Sales (USD)")
  """
  def set_chart_axis_titles(
        %Spreadsheet{reference: ref},
        sheet_name,
        chart_index,
        category_axis_title,
        value_axis_title
      ) do
    case UmyaNative.set_chart_axis_titles(
           ref,
           sheet_name,
           chart_index,
           category_axis_title,
           value_axis_title
         ) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end
end
