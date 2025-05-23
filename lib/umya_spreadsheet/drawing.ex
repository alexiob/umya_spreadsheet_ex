defmodule UmyaSpreadsheet.Drawing do
  @moduledoc """
  Functions for working with shapes, text boxes, and connectors in Excel spreadsheets.
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaNative

  @doc """
  Adds a shape to a spreadsheet at the specified cell position.

  ## Parameters

  * `spreadsheet` - A spreadsheet struct
  * `sheet_name` - The name of the sheet to add the shape to
  * `cell_address` - The cell address where the shape should be placed (e.g., "A1")
  * `shape_type` - The type of shape to add. Supported types:
    * "rectangle" - Rectangle shape
    * "ellipse" or "oval" or "circle" - Ellipse/oval shape
    * "rounded_rectangle" - Rectangle with rounded corners
    * "triangle" - Triangle shape
    * "right_triangle" - Right triangle shape
    * "pentagon" - Pentagon shape
    * "hexagon" - Hexagon shape
    * "octagon" - Octagon shape
    * "trapezoid" - Trapezoid shape
    * "diamond" - Diamond shape
    * "arrow" - Arrow shape
    * "line" - Line shape
    * "connector" - Connector line (use add_connector/6 for connecting two cells)
  * `width` - The width of the shape in pixels
  * `height` - The height of the shape in pixels
  * `fill_color` - The fill color for the shape. Can be a named color (e.g., "red", "blue") or a hex color code (e.g., "FF0000")
  * `outline_color` - The outline/border color for the shape. Can be a named color or hex code
  * `outline_width` - The width of the outline/border in points

  ## Returns

  * `:ok` - Shape was successfully added
  * `{:error, :not_found}` - Sheet was not found
  * `{:error, :error}` - Failed to add shape for another reason

  ## Examples

      # Add a blue rectangle at cell A1
      UmyaSpreadsheet.add_shape(spreadsheet, "Sheet1", "A1", "rectangle", 200, 100, "#0000FF", "black", 1.0)

      # Add a red circle at cell B5
      UmyaSpreadsheet.add_shape(spreadsheet, "Sheet1", "B5", "ellipse", 150, 150, "red", "black", 1.0)
  """
  @spec add_shape(
          Spreadsheet.t(),
          String.t(),
          String.t(),
          String.t(),
          float(),
          float(),
          String.t(),
          String.t(),
          float()
        ) :: :ok | {:error, atom()}
  def add_shape(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_address,
        shape_type,
        width,
        height,
        fill_color,
        outline_color,
        outline_width
      ) do
    case UmyaNative.add_shape(
           ref,
           sheet_name,
           cell_address,
           shape_type,
           width,
           height,
           fill_color,
           outline_color,
           outline_width
         ) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Adds a text box to a spreadsheet at the specified cell position.

  A text box is a rectangular shape with text inside it.

  ## Parameters

  * `spreadsheet` - A spreadsheet struct
  * `sheet_name` - The name of the sheet to add the text box to
  * `cell_address` - The cell address where the text box should be placed (e.g., "A1")
  * `text` - The text to display inside the text box
  * `width` - The width of the text box in pixels
  * `height` - The height of the text box in pixels
  * `fill_color` - The background color for the text box. Can be a named color (e.g., "white") or a hex color code
  * `text_color` - The color of the text. Can be a named color (e.g., "black") or a hex color code
  * `outline_color` - The border color for the text box. Can be a named color or hex code
  * `outline_width` - The width of the border in points

  ## Returns

  * `:ok` - Text box was successfully added
  * `{:error, :not_found}` - Sheet was not found
  * `{:error, :error}` - Failed to add text box for another reason

  ## Examples

      # Add a white text box with black text at cell A1
      UmyaSpreadsheet.add_text_box(
        spreadsheet,
        "Sheet1",
        "A1",
        "Hello World",
        200,
        100,
        "white",
        "black",
        "gray",
        1.0
      )
  """
  @spec add_text_box(
          Spreadsheet.t(),
          String.t(),
          String.t(),
          String.t(),
          float(),
          float(),
          String.t(),
          String.t(),
          String.t(),
          float()
        ) :: :ok | {:error, atom()}
  def add_text_box(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_address,
        text,
        width,
        height,
        fill_color,
        text_color,
        outline_color,
        outline_width
      ) do
        case UmyaNative.add_text_box(
               ref,
               sheet_name,
               cell_address,
               text,
               width,
               height,
               fill_color,
               text_color,
               outline_color,
               outline_width
             ) do
          {:ok, :ok} -> :ok
          :ok -> :ok
          result -> result
        end
  end

  @doc """
  Adds a connector line between two cells in a spreadsheet.

  A connector is a line that connects two cells, useful for creating flowcharts and diagrams.

  ## Parameters

  * `spreadsheet` - A spreadsheet struct
  * `sheet_name` - The name of the sheet to add the connector to
  * `from_cell` - The starting cell address for the connector (e.g., "A1")
  * `to_cell` - The ending cell address for the connector (e.g., "B5")
  * `line_color` - The color of the connector line. Can be a named color (e.g., "black") or a hex color code
  * `line_width` - The width of the connector line in points

  ## Returns

  * `:ok` - Connector was successfully added
  * `{:error, :not_found}` - Sheet was not found
  * `{:error, :error}` - Failed to add connector for another reason

  ## Examples

      # Add a red connector line from cell A1 to cell C3
      UmyaSpreadsheet.add_connector(spreadsheet, "Sheet1", "A1", "C3", "red", 1.5)
  """
  @spec add_connector(
          Spreadsheet.t(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          float()
        ) :: :ok | {:error, atom()}
  def add_connector(
        %Spreadsheet{reference: ref},
        sheet_name,
        from_cell,
        to_cell,
        line_color,
        line_width
      ) do
        case UmyaNative.add_connector(
               ref,
               sheet_name,
               from_cell,
               to_cell,
               line_color,
               line_width
             ) do
          {:ok, :ok} -> :ok
          :ok -> :ok
          result -> result
        end
  end
end
