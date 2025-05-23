# Charts

UmyaSpreadsheet supports the creation of various types of charts in spreadsheets.
You can add charts to your spreadsheet using the `add_chart/9` function or the
more advanced `add_chart_with_options/16` function, which allows for greater customization.

## Adding Charts

### Basic Chart Addition

```elixir
UmyaSpreadsheet.add_chart(
  spreadsheet,
  "Sheet1",
  "LineChart",  # Chart type
  "E1",         # From cell
  "J10",        # To cell
  "Sales Data", # Chart title
  ["Sheet1!$B$2:$B$5", "Sheet1!$C$2:$C$5"],  # Data series
  ["Revenue", "Expenses"],                   # Series titles
  ["Q1", "Q2", "Q3", "Q4"]                   # Point titles
)
```

### Advanced Chart Addition with Options

```elixir
UmyaSpreadsheet.add_chart_with_options(
  spreadsheet,
  "Sheet1",
  "Bar3DChart",  # Chart type
  "E12",         # From cell
  "J22",         # To cell
  "Quarterly Results",  # Chart title
  ["Sheet1!$B$2:$B$5", "Sheet1!$C$2:$C$5"],  # Data series
  ["Revenue", "Expenses"],                   # Series titles
  ["Q1", "Q2", "Q3", "Q4"],                  # Point titles
  1,              # Style (1-48)
  true,           # Vary colors
  %{              # 3D view options
    rot_x: 15,
    rot_y: 20,
    perspective: 30,
    height_percent: 100
  },
  %{              # Legend options
    position: "right",
    overlay: false
  },
  %{},            # Axis options (empty map for defaults)
  %{              # Data label options
    show_values: true,
    show_percent: false
  },
  %{}             # Chart-specific options (empty map for defaults)
)
```

### Parameters

- `spreadsheet`: The spreadsheet object to which the chart will be added.
- `sheet_name`: The name of the sheet where the chart will be placed.
- `chart_type`: The type of chart to create (e.g., `"LineChart"`, `"BarChart"`).
- `from_cell`: The starting cell for the chart (e.g., `"E1"`).
- `to_cell`: The ending cell for the chart (e.g., `"J10"`).
- `title`: The title of the chart.
- `data_series`: A list of data series to be plotted (e.g., `["Sheet1!$B$2:$B$5", "Sheet1!$C$2:$C$5"]`).
- `series_titles`: A list of titles for each data series (e.g., `["Revenue", "Expenses"]`).
- `point_titles`: A list of titles for each point on the chart (e.g., `["Q1", "Q2", "Q3", "Q4"]`).
- `style`: (Optional) The style of the chart (1-48).
- `vary_colors`: (Optional) Boolean indicating whether to vary colors for different series.
- `view_options`: (Optional) A map of 3D view options (e.g., rotation angles, perspective).
- `legend_options`: (Optional) A map of legend options (e.g., position, overlay).
- `axis_options`: (Optional) A map of axis options (e.g., min/max values, titles).
- `data_label_options`: (Optional) A map of data label options (e.g., show values, show percentages).
- `chart_options`: (Optional) A map of chart-specific options (e.g., gridlines, background color).
- `chart_specific_options`: (Optional) A map of additional chart-specific options.

## Supported Chart Types

The following chart types are supported for use with `add_chart/9` and `add_chart_with_options/16`:

| Chart Type        | Description                                        |
| ----------------- | -------------------------------------------------- |
| `"LineChart"`     | Standard 2D line chart                             |
| `"Line3DChart"`   | 3D line chart with perspective                     |
| `"PieChart"`      | Standard 2D pie chart                              |
| `"Pie3DChart"`    | 3D pie chart with perspective                      |
| `"DoughnutChart"` | Doughnut chart (pie with hole in center)           |
| `"ScatterChart"`  | X-Y scatter plot                                   |
| `"BarChart"`      | Standard 2D bar/column chart                       |
| `"Bar3DChart"`    | 3D bar/column chart with perspective               |
| `"RadarChart"`    | Radar/spider chart                                 |
| `"BubbleChart"`   | Bubble chart (scatter with variable point sizes)   |
| `"AreaChart"`     | Standard 2D area chart                             |
| `"Area3DChart"`   | 3D area chart with perspective                     |
| `"OfPieChart"`    | "Of Pie" chart variant (pie with breakout section) |

### Example Usage

```elixir
# Create a basic line chart
UmyaSpreadsheet.add_chart(
  spreadsheet,
  "Sheet1",
  "LineChart",  # Chart type
  "E1",         # From cell
  "J10",        # To cell
  "Sales Data", # Chart title
  ["Sheet1!$B$2:$B$5", "Sheet1!$C$2:$C$5"],  # Data series
  ["Revenue", "Expenses"],                    # Series titles
  ["Q1", "Q2", "Q3", "Q4"]                    # Point titles
)

# Create a 3D bar chart with more options
UmyaSpreadsheet.add_chart_with_options(
  spreadsheet,
  "Sheet1",
  "Bar3DChart",  # Chart type
  "E12",         # From cell
  "J22",         # To cell
  "Quarterly Results",  # Chart title
  ["Sheet1!$B$2:$B$5", "Sheet1!$C$2:$C$5"],  # Data series
  ["Revenue", "Expenses"],                    # Series titles
  ["Q1", "Q2", "Q3", "Q4"],                   # Point titles
  1,              # Style (1-48)
  true,           # Vary colors
  %{              # 3D view options
    rot_x: 15,
    rot_y: 20,
    perspective: 30,
    height_percent: 100
  },
  %{              # Legend options
    position: "right",
    overlay: false
  },
  %{},            # Axis options (empty map for defaults)
  %{              # Data label options
    show_values: true,
    show_percent: false
  },
  %{}             # Chart-specific options (empty map for defaults)
)
```
