defmodule UmyaSpreadsheet.AdvancedFillFunctions do
  @moduledoc """
  Functions for applying advanced fill patterns to cells in a spreadsheet, including gradient fills and pattern fills.

  This module provides functionality for:
  - Creating gradient fills with multiple color stops
  - Applying pattern fills with various patterns
  - Managing complex fill styles beyond simple background colors
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaSpreadsheet.ErrorHandling
  alias UmyaNative

  @doc """
  Sets a gradient fill for a cell with custom gradient stops.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `degree` - The angle of the gradient in degrees (0-360)
  - `gradient_stops` - List of tuples `{position, color}` where position is 0.0-1.0 and color is a hex string

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      # Create a gradient from red to blue
      gradient_stops = [{0.0, "#FF0000"}, {1.0, "#0000FF"}]
      :ok = UmyaSpreadsheet.AdvancedFillFunctions.set_gradient_fill(
        spreadsheet, "Sheet1", "A1", 45.0, gradient_stops
      )
  """
  def set_gradient_fill(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_address,
        degree,
        gradient_stops
      ) do
    UmyaNative.set_gradient_fill(ref, sheet_name, cell_address, degree, gradient_stops)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets a simple linear gradient fill between two colors.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `start_color` - The starting color as a hex string
  - `end_color` - The ending color as a hex string
  - `angle` - Optional angle in degrees (default: 0.0)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      # Create a horizontal gradient from green to yellow
      :ok = UmyaSpreadsheet.AdvancedFillFunctions.set_linear_gradient_fill(
        spreadsheet, "Sheet1", "A1", "#00FF00", "#FFFF00", 0.0
      )
  """
  def set_linear_gradient_fill(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_address,
        start_color,
        end_color,
        angle \\ 0.0
      ) do
    UmyaNative.set_linear_gradient_fill(
      ref,
      sheet_name,
      cell_address,
      start_color,
      end_color,
      angle
    )
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets a radial gradient fill from center to edge.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `center_color` - The center color as a hex string
  - `edge_color` - The edge color as a hex string

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      # Create a radial gradient from white center to blue edge
      :ok = UmyaSpreadsheet.AdvancedFillFunctions.set_radial_gradient_fill(
        spreadsheet, "Sheet1", "A1", "#FFFFFF", "#0000FF"
      )
  """
  def set_radial_gradient_fill(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_address,
        center_color,
        edge_color
      ) do
    UmyaNative.set_radial_gradient_fill(ref, sheet_name, cell_address, center_color, edge_color)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets a three-color gradient fill with start, middle, and end colors.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `start_color` - The starting color as a hex string
  - `middle_color` - The middle color as a hex string
  - `end_color` - The ending color as a hex string
  - `angle` - Optional angle in degrees (default: 0.0)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      # Create a three-color gradient: red -> yellow -> green
      :ok = UmyaSpreadsheet.AdvancedFillFunctions.set_three_color_gradient_fill(
        spreadsheet, "Sheet1", "A1", "#FF0000", "#FFFF00", "#00FF00", 45.0
      )
  """
  def set_three_color_gradient_fill(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_address,
        start_color,
        middle_color,
        end_color,
        angle \\ 0.0
      ) do
    UmyaNative.set_three_color_gradient_fill(
      ref,
      sheet_name,
      cell_address,
      start_color,
      middle_color,
      end_color,
      angle
    )
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets a custom gradient fill with multiple color stops and validation.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `degree` - The angle of the gradient in degrees (0-360)
  - `gradient_stops` - List of tuples `{position, color}` where position is 0.0-1.0
  - `validate_positions` - Whether to validate and sort gradient stop positions (default: true)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      # Create a complex multi-color gradient
      gradient_stops = [
        {0.0, "#FF0000"},   # Red at start
        {0.25, "#FFFF00"},  # Yellow at 25%
        {0.5, "#00FF00"},   # Green at 50%
        {0.75, "#00FFFF"},  # Cyan at 75%
        {1.0, "#0000FF"}    # Blue at end
      ]
      :ok = UmyaSpreadsheet.AdvancedFillFunctions.set_custom_gradient_fill(
        spreadsheet, "Sheet1", "A1", 90.0, gradient_stops, true
      )
  """
  def set_custom_gradient_fill(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_address,
        degree,
        gradient_stops,
        validate_positions \\ true
      ) do
    UmyaNative.set_custom_gradient_fill(
      ref,
      sheet_name,
      cell_address,
      degree,
      gradient_stops,
      validate_positions
    )
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the gradient fill information for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - `{:ok, {degree, gradient_stops}}` on success where gradient_stops is a list of `{position, color}` tuples
  - `{:error, reason}` on failure or if no gradient fill exists

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      case UmyaSpreadsheet.AdvancedFillFunctions.get_gradient_fill(spreadsheet, "Sheet1", "A1") do
        {:ok, {degree, stops}} ->
          IO.inspect("Gradient angle: " <> degree})
          IO.inspect("Gradient stops: " <> inspect(stops))
        {:error, reason} ->
          IO.puts("No gradient fill: " <> reason)
      end
  """
  def get_gradient_fill(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    UmyaNative.get_gradient_fill(ref, sheet_name, cell_address)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets a pattern fill for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `pattern_type` - The pattern type (atom or string)
  - `foreground_color` - The foreground color as a hex string
  - `background_color` - Optional background color as a hex string

  ## Pattern Types

  Available pattern types include:
  - `:solid` - Solid fill
  - `:dark_gray` - Dark gray pattern
  - `:medium_gray` - Medium gray pattern
  - `:light_gray` - Light gray pattern
  - `:gray_125` - 12.5% gray pattern
  - `:gray_0625` - 6.25% gray pattern
  - `:dark_horizontal` - Dark horizontal lines
  - `:dark_vertical` - Dark vertical lines
  - `:dark_down` - Dark diagonal lines (down)
  - `:dark_up` - Dark diagonal lines (up)
  - `:dark_grid` - Dark grid pattern
  - `:dark_trellis` - Dark trellis pattern
  - `:light_horizontal` - Light horizontal lines
  - `:light_vertical` - Light vertical lines
  - `:light_down` - Light diagonal lines (down)
  - `:light_up` - Light diagonal lines (up)
  - `:light_grid` - Light grid pattern
  - `:light_trellis` - Light trellis pattern

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      # Apply a solid red fill
      :ok = UmyaSpreadsheet.AdvancedFillFunctions.set_pattern_fill(
        spreadsheet, "Sheet1", "A1", :solid, "#FF0000"
      )

      # Apply a diagonal pattern with blue foreground and white background
      :ok = UmyaSpreadsheet.AdvancedFillFunctions.set_pattern_fill(
        spreadsheet, "Sheet1", "B1", :dark_down, "#0000FF", "#FFFFFF"
      )
  """
  def set_pattern_fill(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_address,
        pattern_type,
        foreground_color,
        background_color \\ nil
      ) do
    pattern_str = if is_atom(pattern_type), do: Atom.to_string(pattern_type), else: pattern_type

    UmyaNative.set_pattern_fill(
      ref,
      sheet_name,
      cell_address,
      pattern_str,
      foreground_color,
      background_color
    )
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the pattern fill information for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - `{:ok, {pattern_type, foreground_color, background_color}}` on success
  - `{:error, reason}` on failure or if no pattern fill exists

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      case UmyaSpreadsheet.AdvancedFillFunctions.get_pattern_fill(spreadsheet, "Sheet1", "A1") do
        {:ok, {pattern, fg_color, bg_color}} ->
          IO.inspect("Pattern: " <> pattern <> ", FG: " <> fg_color <> ", BG: " <> bg_color)
        {:error, reason} ->
          IO.puts("No pattern fill: " <> reason)
      end
  """
  def get_pattern_fill(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    UmyaNative.get_pattern_fill(ref, sheet_name, cell_address)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Clears all fill formatting from a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")

      # Clear all fill formatting from cell A1
      :ok = UmyaSpreadsheet.AdvancedFillFunctions.clear_fill(spreadsheet, "Sheet1", "A1")
  """
  def clear_fill(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    UmyaNative.clear_fill(ref, sheet_name, cell_address)
    |> ErrorHandling.standardize_result()
  end
end
