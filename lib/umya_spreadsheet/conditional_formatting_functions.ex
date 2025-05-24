defmodule UmyaSpreadsheet.ConditionalFormatting do
  @moduledoc """
  Functions for conditional formatting.
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaNative
  alias UmyaSpreadsheetEx.CustomStructs.CustomColor

  @doc """
  Adds a cell value rule for conditional formatting.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `range` - The cell range to apply formatting to (e.g., "A1:A10")
  - `operator` - The comparison operator ("equal", "notEqual", "greaterThan", "lessThan", "greaterThanOrEqual", "lessThanOrEqual", "between", "notBetween")
  - `value1` - The first value for comparison
  - `value2` - The second value for comparison (required for "between" and "notBetween" operators)
  - `format_style` - The color to apply when the condition is met (e.g., "#FF0000" for red)

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      # Highlight cells equal to 50 in red
      :ok = UmyaSpreadsheet.add_cell_value_rule(
        spreadsheet,
        "Sheet1",
        "A1:A10",
        "equal",
        "50",
        nil,
        "#FF0000"
      )

      # Highlight cells between 30 and 70 in green
      :ok = UmyaSpreadsheet.add_cell_value_rule(
        spreadsheet,
        "Sheet1",
        "A1:A10",
        "between",
        "30",
        "70",
        "#00FF00"
      )
  """
  def add_cell_value_rule(%Spreadsheet{reference: ref}, sheet_name, range, operator, value1, value2, format_style) do
    case UmyaNative.add_cell_value_rule(ref, sheet_name, range, operator, value1, value2, format_style) do
      {:ok, :ok} -> :ok
      {:ok, true} -> :ok  # Handle the true boolean return value
      :ok -> :ok  # Handle simple ok atoms
      {:error, reason} -> {:error, reason}
      result -> result
    end
  end

  @doc """
  Adds a cell "is" rule for conditional formatting based on text conditions.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `range` - The cell range to apply formatting to (e.g., "A1:A10")
  - `operator` - The text comparison operator ("beginsWith", "endsWith", "contains", "notContains")
  - `value1` - The text value for comparison
  - `value2` - Not used, pass nil
  - `format_style` - The color to apply when the condition is met (e.g., "#FF0000" for red)

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      # Highlight cells that begin with "Error" in red
      :ok = UmyaSpreadsheet.add_cell_is_rule(
        spreadsheet,
        "Sheet1",
        "A1:A10",
        "beginsWith",
        "Error",
        nil,
        "#FF0000"
      )
  """
  def add_cell_is_rule(%Spreadsheet{reference: ref}, sheet_name, range, operator, value1, value2, format_style) do
    case UmyaNative.add_cell_is_rule(ref, sheet_name, range, operator, value1, value2, format_style) do
      {:ok, :ok} -> :ok
      result -> result
    end
  end

  @doc """
  Adds a top/bottom rule for conditional formatting.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `range` - The cell range to apply formatting to (e.g., "A1:A10")
  - `rule_type` - The rule type ("top", "bottom")
  - `rank` - The number of top/bottom items to highlight
  - `percent` - Whether rank is a percentage (true) or a count (false)
  - `format_style` - The color to apply when the condition is met (e.g., "#FF0000" for red)

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      # Highlight top 3 values in green
      :ok = UmyaSpreadsheet.add_top_bottom_rule(
        spreadsheet,
        "Sheet1",
        "A1:A10",
        "top",
        3,
        false,
        "#00FF00"
      )

      # Highlight bottom 10% of values in red
      :ok = UmyaSpreadsheet.add_top_bottom_rule(
        spreadsheet,
        "Sheet1",
        "A1:A10",
        "bottom",
        10,
        true,
        "#FF0000"
      )
  """
  def add_top_bottom_rule(%Spreadsheet{reference: ref}, sheet_name, range, rule_type, rank, percent, format_style) do
    case UmyaNative.add_top_bottom_rule(ref, sheet_name, range, rule_type, rank, percent, format_style) do
      {:ok, :ok} -> :ok
      result -> result
    end
  end

  @doc """
  Adds a data bar conditional formatting rule to a range of cells.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `range` - The cell range to apply formatting to (e.g., "A1:A10")
  - `min_value` - Tuple of {type, value} for minimum, or nil for automatic
  - `max_value` - Tuple of {type, value} for maximum, or nil for automatic
  - `color` - The color to use for the data bars (e.g., "#638EC6")

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      # Add data bars with automatic min/max
      :ok = UmyaSpreadsheet.add_data_bar(
        spreadsheet,
        "Sheet1",
        "A1:A10",
        nil,
        nil,
        "#638EC6"
      )

      # Add data bars with fixed min/max
      :ok = UmyaSpreadsheet.add_data_bar(
        spreadsheet,
        "Sheet1",
        "B1:B10",
        {"num", "0"},
        {"num", "100"},
        "#FF0000"
      )
  """
  def add_data_bar(%Spreadsheet{reference: ref}, sheet_name, range, min_value, max_value, color) do
    case UmyaNative.add_data_bar(ref, sheet_name, range, min_value, max_value, color) do
      {:ok, :ok} -> :ok
      result -> result
    end
  end

  @doc """
  Adds a text rule conditional formatting rule.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `range` - The cell range to apply formatting to (e.g., "A1:A10")
  - `operator` - The text operator ("contains", "notContains", "beginsWith", "endsWith")
  - `text` - The text to search for
  - `format_style` - The color to apply when the condition is met (e.g., "#FF0000" for red)

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      # Highlight cells containing "Error" in red
      :ok = UmyaSpreadsheet.add_text_rule(
        spreadsheet,
        "Sheet1",
        "A1:A10",
        "contains",
        "Error",
        "#FF0000"
      )

      # Highlight cells beginning with "Warning" in orange
      :ok = UmyaSpreadsheet.add_text_rule(
        spreadsheet,
        "Sheet1",
        "A1:A10",
        "beginsWith",
        "Warning",
        "#FFA500"
      )
  """
  def add_text_rule(%Spreadsheet{reference: ref}, sheet_name, range, operator, text, format_style) do
    case UmyaNative.add_text_rule(ref, sheet_name, range, operator, text, format_style) do
      {:ok, :ok} -> :ok
      result -> result
    end
  end

  @doc """
  Adds a color scale conditional formatting rule to a range of cells (two-color version).

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `range` - The cell range to apply formatting to (e.g., "A1:A10")
  - `min_type` - The minimum value type-value tuple or nil for automatic
  - `min_value` - The minimum value (if min_type is not nil)
  - `min_color` - The color for minimum values (e.g., "#FF0000" for red)
  - `max_type` - The maximum value type-value tuple or nil for automatic
  - `max_value` - The maximum value (if max_type is not nil)
  - `max_color` - The color for maximum values (e.g., "#00FF00" for green)

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      # Add a color scale with automatic min/max (red to green)
      :ok = UmyaSpreadsheet.add_color_scale(
        spreadsheet,
        "Sheet1",
        "A1:A10",
        nil,
        nil,
        nil,
        "#FF0000",
        nil,
        "#00FF00"
      )
  """
  def add_color_scale(%Spreadsheet{reference: ref}, sheet_name, range, min_type, min_value, min_color, max_type, max_value, max_color) do
    # Process parameters to ensure they're in the right format
    min_type_str = process_type(min_type, "min")
    max_type_str = process_type(max_type, "max")

    min_value_str = process_value(min_value)
    max_value_str = process_value(max_value)

    min_color_map = convert_color(min_color)
    max_color_map = convert_color(max_color)

    # Call the NIF with explicit nil values for mid parameters
    try do
      case UmyaNative.add_color_scale(
        ref,
        sheet_name,
        range,
        min_type_str,
        min_value_str,
        min_color_map,
        nil,  # mid_type
        nil,  # mid_value
        nil,  # mid_color
        max_type_str,
        max_value_str,
        max_color_map
      ) do
        {:ok, :ok} -> :ok
        {:ok, true} -> :ok
        :ok -> :ok
        {:error, reason} -> {:error, reason}
        other ->
          IO.puts("Unexpected response from add_color_scale: #{inspect(other)}")
          :ok
      end
    rescue
      e ->
        IO.puts("Error in add_color_scale: #{inspect(e, pretty: true)}")
        IO.puts("Min color: #{inspect(min_color_map, pretty: true)}")
        IO.puts("Max color: #{inspect(max_color_map, pretty: true)}")
        {:error, "Function call failed: #{Exception.message(e)}"}
    end
  end

  @doc """
  Adds a color scale conditional formatting rule to a range of cells (three-color version).

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `range` - The cell range to apply formatting to (e.g., "A1:A10")
  - `min_type` - The minimum value type string or nil for automatic
  - `min_value` - The minimum value or nil for automatic
  - `min_color` - The color or color map for minimum values (e.g., "#FF0000" for red)
  - `mid_type` - The midpoint value type string or nil for automatic
  - `mid_value` - The midpoint value or nil for automatic
  - `mid_color` - The color or color map for midpoint values (e.g., "#FFFF00" for yellow)
  - `max_type` - The maximum value type string or nil for automatic
  - `max_value` - The maximum value or nil for automatic
  - `max_color` - The color or color map for maximum values (e.g., "#00FF00" for green)

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      # Add a three-color scale (red-yellow-green)
      :ok = UmyaSpreadsheet.add_color_scale(
        spreadsheet,
        "Sheet1",
        "A1:A10",
        "min",
        nil,
        %{argb: "FFFF0000"},
        "percentile",
        "50",
        %{argb: "FFFFFF00"},
        "max",
        nil,
        %{argb: "FF00FF00"}
      )
  """
  def add_color_scale(%Spreadsheet{reference: ref}, sheet_name, range, min_type, min_value, min_color, mid_type, mid_value, mid_color, max_type, max_value, max_color) do
    # Process parameters to ensure they're in the right format
    min_type_str = process_type(min_type, "min")
    mid_type_str = process_type(mid_type, "percentile")
    max_type_str = process_type(max_type, "max")

    min_value_str = process_value(min_value)
    mid_value_str = process_value(mid_value)
    max_value_str = process_value(max_value)

    min_color_map = convert_color(min_color)
    mid_color_map = convert_color(mid_color)
    max_color_map = convert_color(max_color)

    # Call the NIF with explicit error handling
    try do
      case UmyaNative.add_color_scale(
        ref,
        sheet_name,
        range,
        min_type_str,
        min_value_str,
        min_color_map,
        mid_type_str,
        mid_value_str,
        mid_color_map,
        max_type_str,
        max_value_str,
        max_color_map
      ) do
        {:ok, :ok} -> :ok
        {:ok, true} -> :ok
        :ok -> :ok
        {:error, reason} -> {:error, reason}
        other ->
          IO.puts("Unexpected response from add_color_scale (3-color): #{inspect(other)}")
          :ok
      end
    rescue
      e ->
        IO.puts("Error in add_color_scale (3-color): #{inspect(e, pretty: true)}")
        IO.puts("Min color: #{inspect(min_color_map, pretty: true)}")
        IO.puts("Mid color: #{inspect(mid_color_map, pretty: true)}")
        IO.puts("Max color: #{inspect(max_color_map, pretty: true)}")
        {:error, "Function call failed: #{Exception.message(e)}"}
    end
  end

  # New helper functions to simplify parameter processing
  defp process_type(nil, default), do: default
  defp process_type({type, _}, _) when is_binary(type), do: type
  defp process_type(type, _) when is_binary(type), do: type
  defp process_type(_, default), do: default

  defp process_value(nil), do: nil
  defp process_value({_, value}) when is_binary(value), do: if(value == "", do: nil, else: value)
  defp process_value(value) when is_binary(value), do: if(value == "", do: nil, else: value)
  defp process_value(_), do: nil

  # Helper function to convert various color formats to a CustomColor struct with argb field
  alias UmyaSpreadsheetEx.CustomStructs.CustomColor

  defp convert_color(nil), do: nil
  defp convert_color(%CustomColor{} = color), do: color  # Already a proper CustomColor struct
  defp convert_color(%{argb: argb}) when is_binary(argb), do: %CustomColor{argb: argb}  # Map with argb field
  defp convert_color(hex_color) when is_binary(hex_color) do
    # Use the CustomColor.from_hex function to create a proper struct
    CustomColor.from_hex(hex_color)
  end
  defp convert_color(_), do: %CustomColor{argb: "FFFFFFFF"}  # Default to white for unknown formats

  @doc """
  Adds an icon set rule for conditional formatting.

  Icon sets display different icons based on the values in the cells, providing a visual
  representation of data trends.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `range` - The cell range to apply formatting to (e.g., "A1:A10")
  - `icon_style` - The style of icons to use (currently supports basic icon sets)
  - `thresholds` - A list of threshold tuples {type, value} defining the icon boundaries

  ## Threshold Types

  - `{"min", ""}` - Use the minimum value in the range
  - `{"max", ""}` - Use the maximum value in the range
  - `{"number", "X"}` - Use a specific number
  - `{"percent", "X"}` - Use X percent of the range
  - `{"percentile", "X"}` - Use the Xth percentile
  - `{"formula", "X"}` - Use a formula result

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      # Simple 3-icon set with percentile thresholds
      :ok = UmyaSpreadsheet.add_icon_set(
        spreadsheet,
        "Sheet1",
        "A1:A10",
        "3_traffic_lights",
        [
          {"percentile", "33"},
          {"percentile", "67"}
        ]
      )

      # 5-icon set with custom number thresholds
      :ok = UmyaSpreadsheet.add_icon_set(
        spreadsheet,
        "Sheet1",
        "B1:B10",
        "5_arrows",
        [
          {"number", "20"},
          {"number", "40"},
          {"number", "60"},
          {"number", "80"}
        ]
      )
  """
  def add_icon_set(%Spreadsheet{reference: ref}, sheet_name, range, icon_style, thresholds) do
    case UmyaNative.add_icon_set(ref, sheet_name, range, icon_style, thresholds) do
      {:ok, :ok} -> :ok
      {:ok, true} -> :ok  # Handle the true boolean return value
      :ok -> :ok  # Handle simple ok atoms
      {:error, reason} -> {:error, reason}
      result -> result
    end
  end

  @doc """
  Adds an above or below average rule for conditional formatting.

  This rule highlights cells that are above or below the average value of the selected range.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `range` - The cell range to apply formatting to (e.g., "A1:A10")
  - `rule_type` - The type of rule: "above", "below", "above_equal", "below_equal"
  - `std_dev` - Optional standard deviation for more advanced rules (nil for basic average)
  - `format_style` - The color to apply when the condition is met (e.g., "#FF0000" for red)

  ## Rule Types

  - `"above"` - Highlight cells above the average
  - `"below"` - Highlight cells below the average
  - `"above_equal"` - Highlight cells above or equal to the average
  - `"below_equal"` - Highlight cells below or equal to the average

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      # Highlight cells above average in green
      :ok = UmyaSpreadsheet.add_above_below_average_rule(
        spreadsheet,
        "Sheet1",
        "A1:A10",
        "above",
        nil,
        "#00FF00"
      )

      # Highlight cells below average in red
      :ok = UmyaSpreadsheet.add_above_below_average_rule(
        spreadsheet,
        "Sheet1",
        "B1:B10",
        "below",
        nil,
        "#FF0000"
      )

      # Highlight cells above average plus one standard deviation in blue
      :ok = UmyaSpreadsheet.add_above_below_average_rule(
        spreadsheet,
        "Sheet1",
        "C1:C10",
        "above",
        1,
        "#0000FF"
      )
  """
  def add_above_below_average_rule(%Spreadsheet{reference: ref}, sheet_name, range, rule_type, std_dev, format_style) do
    case UmyaNative.add_above_below_average_rule(ref, sheet_name, range, rule_type, std_dev, format_style) do
      {:ok, :ok} -> :ok
      {:ok, true} -> :ok  # Handle the true boolean return value
      :ok -> :ok  # Handle simple ok atoms
      {:error, reason} -> {:error, reason}
      result -> result
    end
  end
end
