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
  def add_cell_value_rule(
        %Spreadsheet{reference: ref},
        sheet_name,
        range,
        operator,
        value1,
        value2,
        format_style
      ) do
    case UmyaNative.add_cell_value_rule(
           ref,
           sheet_name,
           range,
           operator,
           value1,
           value2,
           format_style
         ) do
      {:ok, :ok} -> :ok
      # Handle the true boolean return value
      {:ok, true} -> :ok
      # Handle simple ok atoms
      :ok -> :ok
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
  def add_cell_is_rule(
        %Spreadsheet{reference: ref},
        sheet_name,
        range,
        operator,
        value1,
        value2,
        format_style
      ) do
    case UmyaNative.add_cell_is_rule(
           ref,
           sheet_name,
           range,
           operator,
           value1,
           value2,
           format_style
         ) do
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
  def add_top_bottom_rule(
        %Spreadsheet{reference: ref},
        sheet_name,
        range,
        rule_type,
        rank,
        percent,
        format_style
      ) do
    case UmyaNative.add_top_bottom_rule(
           ref,
           sheet_name,
           range,
           rule_type,
           rank,
           percent,
           format_style
         ) do
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
  - `max_type` - The maximum value type-value tuple or nil for automatic
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
  def add_color_scale(
        %Spreadsheet{reference: ref},
        sheet_name,
        range,
        min_type,
        min_value,
        min_color,
        max_type,
        max_value,
        max_color
      ) do
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
             # mid_type
             nil,
             # mid_value
             nil,
             # mid_color
             nil,
             max_type_str,
             max_value_str,
             max_color_map
           ) do
        {:ok, :ok} ->
          :ok

        {:ok, true} ->
          :ok

        :ok ->
          :ok

        {:error, reason} ->
          {:error, reason}

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
  def add_color_scale(
        %Spreadsheet{reference: ref},
        sheet_name,
        range,
        min_type,
        min_value,
        min_color,
        mid_type,
        mid_value,
        mid_color,
        max_type,
        max_value,
        max_color
      ) do
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
        {:ok, :ok} ->
          :ok

        {:ok, true} ->
          :ok

        :ok ->
          :ok

        {:error, reason} ->
          {:error, reason}

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
  # Already a proper CustomColor struct
  defp convert_color(%CustomColor{} = color), do: color
  # Map with argb field
  defp convert_color(%{argb: argb}) when is_binary(argb), do: %CustomColor{argb: argb}

  defp convert_color(hex_color) when is_binary(hex_color) do
    # Use the CustomColor.from_hex function to create a proper struct
    CustomColor.from_hex(hex_color)
  end

  # Default to white for unknown formats
  defp convert_color(_), do: %CustomColor{argb: "FFFFFFFF"}

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
      # Handle the true boolean return value
      {:ok, true} -> :ok
      # Handle simple ok atoms
      :ok -> :ok
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
  def add_above_below_average_rule(
        %Spreadsheet{reference: ref},
        sheet_name,
        range,
        rule_type,
        std_dev,
        format_style
      ) do
    case UmyaNative.add_above_below_average_rule(
           ref,
           sheet_name,
           range,
           rule_type,
           std_dev,
           format_style
         ) do
      {:ok, :ok} -> :ok
      # Handle the true boolean return value
      {:ok, true} -> :ok
      # Handle simple ok atoms
      :ok -> :ok
      {:error, reason} -> {:error, reason}
      result -> result
    end
  end

  @doc """
  Gets all conditional formatting rules for a sheet or specific range.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `range` - Optional. The cell range to get rules for. If nil, returns all rules for the sheet.

  ## Returns

  A list of maps, each representing a conditional formatting rule with the following keys:

  - `:range` - The cell range the rule applies to
  - `:rule_type` - The type of rule (":cell_is", ":color_scale", ":data_bar", ":icon_set", etc.)
  - Other rule-specific fields depending on the rule type

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      # Get all conditional formatting rules in Sheet1
      rules = UmyaSpreadsheet.get_conditional_formatting_rules(
        spreadsheet,
        "Sheet1"
      )

      # Get conditional formatting rules for a specific range
      range_rules = UmyaSpreadsheet.get_conditional_formatting_rules(
        spreadsheet,
        "Sheet1",
        "A1:A10"
      )
  """
  def get_conditional_formatting_rules(%Spreadsheet{reference: ref}, sheet_name, range \\ nil) do
    case UmyaNative.get_conditional_formatting_rules(ref, sheet_name, range) do
      {:ok, rules} -> rules
      {:error, reason} -> {:error, reason}
      result -> result
    end
  end

  @doc """
  Gets all cell value rules for a sheet or specific range.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `range` - Optional. The cell range to get rules for. If nil, returns all rules for the sheet.

  ## Returns

  A list of maps, each representing a cell value rule with the following keys:

  - `:range` - The cell range the rule applies to
  - `:operator` - The comparison operator (e.g., ":equal", ":greater_than")
  - `:formula` - The formula or value to compare against
  - `:format_style` - The color or style to apply when the condition is met

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      # Get all cell value rules in Sheet1
      rules = UmyaSpreadsheet.get_cell_value_rules(
        spreadsheet,
        "Sheet1"
      )
  """
  def get_cell_value_rules(%Spreadsheet{reference: ref}, sheet_name, range \\ nil) do
    case UmyaNative.get_cell_value_rules(ref, sheet_name, range) do
      {:ok, rules} -> rules
      {:error, reason} -> {:error, reason}
      result -> result
    end
  end

  @doc """
  Gets all color scale rules for a sheet or specific range.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `range` - Optional. The cell range to get rules for. If nil, returns all rules for the sheet.

  ## Returns

  A list of maps, each representing a color scale rule with the following keys:

  - `:range` - The cell range the rule applies to
  - `:min_type` - The type for minimum value (e.g., ":min", ":number", ":percent")
  - `:min_value` - The minimum value (if applicable)
  - `:min_color` - The color for minimum values
  - `:mid_type` - Optional. The type for midpoint value
  - `:mid_value` - Optional. The midpoint value
  - `:mid_color` - Optional. The color for midpoint values
  - `:max_type` - The type for maximum value
  - `:max_value` - The maximum value (if applicable)
  - `:max_color` - The color for maximum values

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      # Get all color scale rules in Sheet1
      rules = UmyaSpreadsheet.get_color_scales(
        spreadsheet,
        "Sheet1"
      )
  """
  def get_color_scales(%Spreadsheet{reference: ref}, sheet_name, range \\ nil) do
    case UmyaNative.get_color_scales(ref, sheet_name, range) do
      {:ok, rules} -> rules
      {:error, reason} -> {:error, reason}
      result -> result
    end
  end

  @doc """
  Gets all data bar rules for a sheet or specific range.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `range` - Optional. The cell range to get rules for. If nil, returns all rules for the sheet.

  ## Returns

  A list of maps, each representing a data bar rule with the following keys:

  - `:range` - The cell range the rule applies to
  - `:min_value` - Optional tuple of {type, value} for minimum
  - `:max_value` - Optional tuple of {type, value} for maximum
  - `:color` - The color of the data bars

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      # Get all data bar rules in Sheet1
      rules = UmyaSpreadsheet.get_data_bars(
        spreadsheet,
        "Sheet1"
      )
  """
  def get_data_bars(%Spreadsheet{reference: ref}, sheet_name, range \\ nil) do
    case UmyaNative.get_data_bars(ref, sheet_name, range) do
      {:ok, rules} -> rules
      {:error, reason} -> {:error, reason}
      result -> result
    end
  end

  @doc """
  Gets all icon set rules for a sheet or specific range.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `range` - Optional. The cell range to get rules for. If nil, returns all rules for the sheet.

  ## Returns

  A list of maps, each representing an icon set rule with the following keys:

  - `:range` - The cell range the rule applies to
  - `:icon_style` - The style of icons to use
  - `:thresholds` - A list of threshold tuples {type, value} defining the icon boundaries

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      # Get all icon set rules in Sheet1
      rules = UmyaSpreadsheet.get_icon_sets(
        spreadsheet,
        "Sheet1"
      )
  """
  def get_icon_sets(%Spreadsheet{reference: ref}, sheet_name, range \\ nil) do
    case UmyaNative.get_icon_sets(ref, sheet_name, range) do
      {:ok, rules} -> rules
      {:error, reason} -> {:error, reason}
      result -> result
    end
  end

  @doc """
  Gets all top/bottom rules for a sheet or specific range.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `range` - Optional. The cell range to get rules for. If nil, returns all rules for the sheet.

  ## Returns

  A list of maps, each representing a top/bottom rule with the following keys:

  - `:range` - The cell range the rule applies to
  - `:rule_type_value` - "top" or "bottom"
  - `:rank` - The number of top/bottom items to highlight
  - `:percent` - Whether rank is a percentage (true) or a count (false)
  - `:format_style` - The color to apply when the condition is met

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      # Get all top/bottom rules in Sheet1
      rules = UmyaSpreadsheet.get_top_bottom_rules(
        spreadsheet,
        "Sheet1"
      )
  """
  def get_top_bottom_rules(%Spreadsheet{reference: ref}, sheet_name, range \\ nil) do
    case UmyaNative.get_top_bottom_rules(ref, sheet_name, range) do
      {:ok, rules} -> rules
      {:error, reason} -> {:error, reason}
      result -> result
    end
  end

  @doc """
  Gets all above/below average rules for a sheet or specific range.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `range` - Optional. The cell range to get rules for. If nil, returns all rules for the sheet.

  ## Returns

  A list of maps, each representing an above/below average rule with the following keys:

  - `:range` - The cell range the rule applies to
  - `:rule_type_value` - "above", "below", "above_equal", or "below_equal"
  - `:std_dev` - Optional standard deviation value
  - `:format_style` - The color to apply when the condition is met

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      # Get all above/below average rules in Sheet1
      rules = UmyaSpreadsheet.get_above_below_average_rules(
        spreadsheet,
        "Sheet1"
      )
  """
  def get_above_below_average_rules(%Spreadsheet{reference: ref}, sheet_name, range \\ nil) do
    case UmyaNative.get_above_below_average_rules(ref, sheet_name, range) do
      {:ok, rules} -> rules
      {:error, reason} -> {:error, reason}
      result -> result
    end
  end

  @doc """
  Gets all text rules for a sheet or specific range.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `range` - Optional. The cell range to get rules for. If nil, returns all rules for the sheet.

  ## Returns

  A list of maps, each representing a text rule with the following keys:

  - `:range` - The cell range the rule applies to
  - `:operator` - The text operator ("contains", "notContains", "beginsWith", "endsWith")
  - `:text` - The text to search for
  - `:format_style` - The color to apply when the condition is met

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      # Get all text rules in Sheet1
      rules = UmyaSpreadsheet.get_text_rules(
        spreadsheet,
        "Sheet1"
      )
  """
  def get_text_rules(%Spreadsheet{reference: ref}, sheet_name, range \\ nil) do
    case UmyaNative.get_text_rules(ref, sheet_name, range) do
      {:ok, rules} -> rules
      {:error, reason} -> {:error, reason}
      result -> result
    end
  end
end
