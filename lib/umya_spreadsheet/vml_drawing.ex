defmodule UmyaSpreadsheet.VmlDrawing do
  @moduledoc """
  Functions for working with VML (Vector Markup Language) drawings in Excel files.

  VML is a legacy format used in Excel for shapes, comments, and embedded objects.
  This module allows you to create and manipulate VML shapes in Excel files.

  ## Examples

      iex> {:ok, spreadsheet} = UmyaSpreadsheet.new()
      iex> UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "Sheet1", "shape1")
      :ok
      iex> UmyaSpreadsheet.VmlDrawing.set_shape_style(spreadsheet, "Sheet1", "shape1", "position:absolute;left:100pt;top:150pt;width:200pt;height:100pt")
      :ok
  """

  alias UmyaSpreadsheet.Spreadsheet

  @doc """
  Creates a new VML shape in the specified worksheet.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - The name of the worksheet
    * `shape_id` - Unique identifier for the shape

  ## Returns

    * `:ok` on success
    * `{:error, reason}` on failure
  """
  @spec create_shape(Spreadsheet.t(), String.t(), String.t()) :: :ok | {:error, String.t()}
  def create_shape(%Spreadsheet{reference: ref}, sheet_name, shape_id) do
    # For now, we're just forwarding to the NIF stub - in the future, this will create a shape
    case UmyaNative.create_vml_shape(UmyaSpreadsheet.unwrap_ref(ref), sheet_name, shape_id) do
      :ok -> :ok
      {:error, reason} -> {:error, "Failed to create VML shape: #{inspect(reason)}"}
    end
  end

  @doc """
  Sets the CSS style for a VML shape.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - The name of the worksheet
    * `shape_id` - Unique identifier for the shape
    * `style` - CSS style string (e.g. "position:absolute;left:100pt;top:150pt;width:200pt;height:100pt")

  ## Returns

    * `:ok` on success
    * `{:error, reason}` on failure
  """
  @spec set_shape_style(Spreadsheet.t(), String.t(), String.t(), String.t()) ::
          :ok | {:error, String.t()}
  def set_shape_style(%Spreadsheet{reference: ref}, sheet_name, shape_id, style) do
    case UmyaNative.set_vml_shape_style(
           UmyaSpreadsheet.unwrap_ref(ref),
           sheet_name,
           shape_id,
           style
         ) do
      :ok -> :ok
      {:error, reason} -> {:error, "Failed to set VML shape style: #{inspect(reason)}"}
    end
  end

  @doc """
  Sets the type of a VML shape.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - The name of the worksheet
    * `shape_id` - Unique identifier for the shape
    * `shape_type` - The shape type (e.g. "rect", "oval", "line", etc.)

  ## Returns

    * `:ok` on success
    * `{:error, reason}` on failure

  ## Shape Types

  Common shape types include:

  * `"rect"` - Rectangle
  * `"oval"` - Oval or circle
  * `"line"` - Line
  * `"polyline"` - Polyline
  * `"roundrect"` - Rounded rectangle
  * `"arc"` - Arc
  """
  @spec set_shape_type(Spreadsheet.t(), String.t(), String.t(), String.t()) ::
          :ok | {:error, String.t()}
  def set_shape_type(%Spreadsheet{reference: ref}, sheet_name, shape_id, shape_type) do
    case UmyaNative.set_vml_shape_type(
           UmyaSpreadsheet.unwrap_ref(ref),
           sheet_name,
           shape_id,
           shape_type
         ) do
      :ok -> :ok
      {:error, reason} -> {:error, "Failed to set VML shape type: #{inspect(reason)}"}
    end
  end

  @doc """
  Sets whether a VML shape is filled.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - The name of the worksheet
    * `shape_id` - Unique identifier for the shape
    * `filled` - Whether the shape should be filled (boolean)

  ## Returns

    * `:ok` on success
    * `{:error, reason}` on failure
  """
  @spec set_shape_filled(Spreadsheet.t(), String.t(), String.t(), boolean()) ::
          :ok | {:error, String.t()}
  def set_shape_filled(%Spreadsheet{reference: ref}, sheet_name, shape_id, filled) do
    case UmyaNative.set_vml_shape_filled(
           UmyaSpreadsheet.unwrap_ref(ref),
           sheet_name,
           shape_id,
           filled
         ) do
      :ok -> :ok
      {:error, reason} -> {:error, "Failed to set VML shape filled property: #{inspect(reason)}"}
    end
  end

  @doc """
  Sets the fill color of a VML shape.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - The name of the worksheet
    * `shape_id` - Unique identifier for the shape
    * `fill_color` - The fill color (e.g. "#FF0000" for red)

  ## Returns

    * `:ok` on success
    * `{:error, reason}` on failure
  """
  @spec set_shape_fill_color(Spreadsheet.t(), String.t(), String.t(), String.t()) ::
          :ok | {:error, String.t()}
  def set_shape_fill_color(%Spreadsheet{reference: ref}, sheet_name, shape_id, fill_color) do
    case UmyaNative.set_vml_shape_fill_color(
           UmyaSpreadsheet.unwrap_ref(ref),
           sheet_name,
           shape_id,
           fill_color
         ) do
      :ok -> :ok
      {:error, reason} -> {:error, "Failed to set VML shape fill color: #{inspect(reason)}"}
    end
  end

  @doc """
  Sets whether a VML shape has a stroke (outline).

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - The name of the worksheet
    * `shape_id` - Unique identifier for the shape
    * `stroked` - Whether the shape should have a stroke (boolean)

  ## Returns

    * `:ok` on success
    * `{:error, reason}` on failure
  """
  @spec set_shape_stroked(Spreadsheet.t(), String.t(), String.t(), boolean()) ::
          :ok | {:error, String.t()}
  def set_shape_stroked(%Spreadsheet{reference: ref}, sheet_name, shape_id, stroked) do
    case UmyaNative.set_vml_shape_stroked(
           UmyaSpreadsheet.unwrap_ref(ref),
           sheet_name,
           shape_id,
           stroked
         ) do
      :ok -> :ok
      {:error, reason} -> {:error, "Failed to set VML shape stroked property: #{inspect(reason)}"}
    end
  end

  @doc """
  Sets the stroke (outline) color of a VML shape.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - The name of the worksheet
    * `shape_id` - Unique identifier for the shape
    * `stroke_color` - The stroke color (e.g. "#0000FF" for blue)

  ## Returns

    * `:ok` on success
    * `{:error, reason}` on failure
  """
  @spec set_shape_stroke_color(Spreadsheet.t(), String.t(), String.t(), String.t()) ::
          :ok | {:error, String.t()}
  def set_shape_stroke_color(%Spreadsheet{reference: ref}, sheet_name, shape_id, stroke_color) do
    case UmyaNative.set_vml_shape_stroke_color(
           UmyaSpreadsheet.unwrap_ref(ref),
           sheet_name,
           shape_id,
           stroke_color
         ) do
      :ok -> :ok
      {:error, reason} -> {:error, "Failed to set VML shape stroke color: #{inspect(reason)}"}
    end
  end

  @doc """
  Sets the stroke (outline) weight/thickness of a VML shape.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - The name of the worksheet
    * `shape_id` - Unique identifier for the shape
    * `stroke_weight` - The stroke weight (e.g. "2pt")

  ## Returns

    * `:ok` on success
    * `{:error, reason}` on failure
  """
  @spec set_shape_stroke_weight(Spreadsheet.t(), String.t(), String.t(), String.t()) ::
          :ok | {:error, String.t()}
  def set_shape_stroke_weight(%Spreadsheet{reference: ref}, sheet_name, shape_id, stroke_weight) do
    case UmyaNative.set_vml_shape_stroke_weight(
           UmyaSpreadsheet.unwrap_ref(ref),
           sheet_name,
           shape_id,
           stroke_weight
         ) do
      :ok -> :ok
      {:error, reason} -> {:error, "Failed to set VML shape stroke weight: #{inspect(reason)}"}
    end
  end
end
