defmodule UmyaNative.VmlSupport do
  @moduledoc false

  # NIF stub functions with appropriate error fallback
  @spec create_vml_shape(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def create_vml_shape(_spreadsheet, _sheet_name, _shape_id),
    do: :erlang.nif_error(:nif_not_loaded)

  @spec set_vml_shape_style(reference(), String.t(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def set_vml_shape_style(_spreadsheet, _sheet_name, _shape_id, _style),
    do: :erlang.nif_error(:nif_not_loaded)

  @spec set_vml_shape_type(reference(), String.t(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def set_vml_shape_type(_spreadsheet, _sheet_name, _shape_id, _shape_type),
    do: :erlang.nif_error(:nif_not_loaded)

  @spec set_vml_shape_filled(reference(), String.t(), String.t(), boolean()) ::
          :ok | {:error, atom()}
  def set_vml_shape_filled(_spreadsheet, _sheet_name, _shape_id, _filled),
    do: :erlang.nif_error(:nif_not_loaded)

  @spec set_vml_shape_fill_color(reference(), String.t(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def set_vml_shape_fill_color(_spreadsheet, _sheet_name, _shape_id, _fill_color),
    do: :erlang.nif_error(:nif_not_loaded)

  @spec set_vml_shape_stroked(reference(), String.t(), String.t(), boolean()) ::
          :ok | {:error, atom()}
  def set_vml_shape_stroked(_spreadsheet, _sheet_name, _shape_id, _stroked),
    do: :erlang.nif_error(:nif_not_loaded)

  @spec set_vml_shape_stroke_color(reference(), String.t(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def set_vml_shape_stroke_color(_spreadsheet, _sheet_name, _shape_id, _stroke_color),
    do: :erlang.nif_error(:nif_not_loaded)

  @spec set_vml_shape_stroke_weight(reference(), String.t(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def set_vml_shape_stroke_weight(_spreadsheet, _sheet_name, _shape_id, _stroke_weight),
    do: :erlang.nif_error(:nif_not_loaded)

  # Getter functions for VML shape properties
  @spec get_vml_shape_style(reference(), String.t(), String.t()) ::
          {:ok, String.t()} | {:error, atom()}
  def get_vml_shape_style(_spreadsheet, _sheet_name, _shape_id),
    do: :erlang.nif_error(:nif_not_loaded)

  @spec get_vml_shape_type(reference(), String.t(), String.t()) ::
          {:ok, String.t()} | {:error, atom()}
  def get_vml_shape_type(_spreadsheet, _sheet_name, _shape_id),
    do: :erlang.nif_error(:nif_not_loaded)

  @spec get_vml_shape_filled(reference(), String.t(), String.t()) ::
          {:ok, boolean()} | {:error, atom()}
  def get_vml_shape_filled(_spreadsheet, _sheet_name, _shape_id),
    do: :erlang.nif_error(:nif_not_loaded)

  @spec get_vml_shape_fill_color(reference(), String.t(), String.t()) ::
          {:ok, String.t()} | {:error, atom()}
  def get_vml_shape_fill_color(_spreadsheet, _sheet_name, _shape_id),
    do: :erlang.nif_error(:nif_not_loaded)

  @spec get_vml_shape_stroked(reference(), String.t(), String.t()) ::
          {:ok, boolean()} | {:error, atom()}
  def get_vml_shape_stroked(_spreadsheet, _sheet_name, _shape_id),
    do: :erlang.nif_error(:nif_not_loaded)

  @spec get_vml_shape_stroke_color(reference(), String.t(), String.t()) ::
          {:ok, String.t()} | {:error, atom()}
  def get_vml_shape_stroke_color(_spreadsheet, _sheet_name, _shape_id),
    do: :erlang.nif_error(:nif_not_loaded)

  @spec get_vml_shape_stroke_weight(reference(), String.t(), String.t()) ::
          {:ok, String.t()} | {:error, atom()}
  def get_vml_shape_stroke_weight(_spreadsheet, _sheet_name, _shape_id),
    do: :erlang.nif_error(:nif_not_loaded)
end
