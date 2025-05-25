defmodule UmyaNative do
  @moduledoc false

  mix_config = Mix.Project.config()
  version = mix_config[:version]
  github_url = mix_config[:package][:links]["GitHub"]

  use RustlerPrecompiled,
    otp_app: :umya_spreadsheet_ex,
    crate: "umya_native",
    version: version,
    base_url: "#{github_url}/releases/download/v#{version}",
    force_build: System.get_env("UMYA_SPREADSHEET_BUILD") in ["1", "true"]

  # Spreadsheet operations
  @spec new_file() :: reference()
  def new_file(), do: error()

  @spec new_file_empty_worksheet() :: reference()
  def new_file_empty_worksheet(), do: error()

  @spec read_file(String.t()) :: reference() | {:error, atom()}
  def read_file(_path), do: error()

  @spec lazy_read_file(String.t()) :: reference() | {:error, atom()}
  def lazy_read_file(_path), do: error()

  @spec write_file(reference(), String.t()) :: :ok | {:error, atom()}
  def write_file(_spreadsheet, _path), do: error()

  @spec write_file_light(reference(), String.t()) :: :ok | {:error, atom()}
  def write_file_light(_spreadsheet, _path), do: error()

  @spec write_file_with_password(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def write_file_with_password(_spreadsheet, _path, _password), do: error()

  @spec write_file_with_password_light(reference(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def write_file_with_password_light(_spreadsheet, _path, _password), do: error()

  # File format options
  @spec write_with_compression(reference(), String.t(), non_neg_integer()) :: :ok | {:error, atom()}
  def write_with_compression(_spreadsheet, _path, _compression_level), do: error()

  @spec write_with_encryption_options(reference(), String.t(), String.t(), String.t(), String.t() | nil, non_neg_integer() | nil) :: :ok | {:error, atom()}
  def write_with_encryption_options(_spreadsheet, _path, _password, _algorithm, _salt_value \\ nil, _spin_count \\ nil), do: error()

  @spec to_binary_xlsx(reference()) :: binary() | {:error, atom()}
  def to_binary_xlsx(_spreadsheet), do: error()

  # CSV export functions
  @spec write_csv(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def write_csv(_spreadsheet, _sheet_name, _path), do: error()

  @spec write_csv_with_options(reference(), String.t(), String.t(), String.t(), String.t(), boolean(), String.t()) :: :ok | {:error, atom()}
  def write_csv_with_options(_spreadsheet, _sheet_name, _path, _encoding, _delimiter, _do_trim, _wrap_with_char), do: error()

  # Aliased functions for compatibility between Rust and Elixir naming
  @spec write_light(reference(), String.t()) :: :ok | {:error, atom()}
  def write_light(spreadsheet, path), do: write_file_light(spreadsheet, path)

  @spec write_with_password_light(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def write_with_password_light(spreadsheet, path, password), do: write_file_with_password_light(spreadsheet, path, password)

  # Cell operations
  @spec get_cell_value(reference(), String.t(), String.t()) :: String.t() | {:error, atom()}
  def get_cell_value(_spreadsheet, _sheet_name, _cell_address), do: error()

  @spec get_formatted_value(reference(), String.t(), String.t()) :: String.t() | {:error, atom()}
  def get_formatted_value(_spreadsheet, _sheet_name, _cell_address), do: error()

  @spec set_cell_value(reference(), String.t(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_cell_value(_spreadsheet, _sheet_name, _cell_address, _value), do: error()

  @spec remove_cell(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def remove_cell(_spreadsheet, _sheet_name, _cell_address), do: error()

  # Sheet operations
  @spec get_sheet_names(reference()) :: [String.t()]
  def get_sheet_names(_spreadsheet), do: error()

  @spec add_sheet(reference(), String.t()) :: :ok | {:error, atom()}
  def add_sheet(_spreadsheet, _sheet_name), do: error()

  @spec clone_sheet(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def clone_sheet(_spreadsheet, _source_sheet_name, _new_sheet_name), do: error()

  @spec remove_sheet(reference(), String.t()) :: :ok | {:error, atom()}
  def remove_sheet(_spreadsheet, _sheet_name), do: error()

  @spec rename_sheet(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def rename_sheet(_spreadsheet, _old_sheet_name, _new_sheet_name), do: error()

  @spec insert_new_row(reference(), String.t(), integer(), integer()) :: :ok | {:error, atom()}
  def insert_new_row(_spreadsheet, _sheet_name, _row_index, _amount), do: error()

  @spec insert_new_column(reference(), String.t(), String.t(), integer()) ::
          :ok | {:error, atom()}
  def insert_new_column(_spreadsheet, _sheet_name, _column, _amount), do: error()

  @spec insert_new_column_by_index(reference(), String.t(), integer(), integer()) ::
          :ok | {:error, atom()}
  def insert_new_column_by_index(_spreadsheet, _sheet_name, _column_index, _amount), do: error()

  @spec remove_row(reference(), String.t(), integer(), integer()) :: :ok | {:error, atom()}
  def remove_row(_spreadsheet, _sheet_name, _row_index, _amount), do: error()

  @spec remove_column(reference(), String.t(), String.t(), integer()) :: :ok | {:error, atom()}
  def remove_column(_spreadsheet, _sheet_name, _column, _amount), do: error()

  @spec remove_column_by_index(reference(), String.t(), integer(), integer()) ::
          :ok | {:error, atom()}
  def remove_column_by_index(_spreadsheet, _sheet_name, _column_index, _amount), do: error()

  @spec move_range(reference(), String.t(), String.t(), integer(), integer()) ::
          :ok | {:error, atom()}
  def move_range(_spreadsheet, _sheet_name, _range, _row, _column), do: error()

  # Style operations
  @spec set_background_color(reference(), String.t(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def set_background_color(_spreadsheet, _sheet_name, _cell_address, _color), do: error()

  @spec set_number_format(reference(), String.t(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def set_number_format(_spreadsheet, _sheet_name, _cell_address, _format_code), do: error()

  @spec set_row_height(reference(), String.t(), integer(), float()) :: :ok | {:error, atom()}
  def set_row_height(_spreadsheet, _sheet_name, _row_number, _height), do: error()

  @spec set_row_style(reference(), String.t(), integer(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def set_row_style(_spreadsheet, _sheet_name, _row_number, _bg_color, _font_color), do: error()

  # Font operations
  @spec set_font_color(reference(), String.t(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_font_color(_spreadsheet, _sheet_name, _cell_address, _color), do: error()

  @spec set_font_size(reference(), String.t(), String.t(), integer()) :: :ok | {:error, atom()}
  def set_font_size(_spreadsheet, _sheet_name, _cell_address, _size), do: error()

  @spec set_font_bold(reference(), String.t(), String.t(), boolean()) :: :ok | {:error, atom()}
  def set_font_bold(_spreadsheet, _sheet_name, _cell_address, _is_bold), do: error()

  @spec set_font_name(reference(), String.t(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_font_name(_spreadsheet, _sheet_name, _cell_address, _font_name), do: error()

  # Cell formatting functions
  @spec set_font_italic(reference(), String.t(), String.t(), boolean()) :: :ok | {:error, atom()}
  def set_font_italic(_spreadsheet, _sheet_name, _cell_address, _is_italic), do: error()

  @spec set_font_underline(reference(), String.t(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_font_underline(_spreadsheet, _sheet_name, _cell_address, _underline_style), do: error()

  @spec set_font_strikethrough(reference(), String.t(), String.t(), boolean()) :: :ok | {:error, atom()}
  def set_font_strikethrough(_spreadsheet, _sheet_name, _cell_address, _is_strikethrough), do: error()

  @spec set_border_style(reference(), String.t(), String.t(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_border_style(_spreadsheet, _sheet_name, _cell_address, _border_position, _border_style), do: error()

  @spec set_cell_rotation(reference(), String.t(), String.t(), integer()) :: :ok | {:error, atom()}
  def set_cell_rotation(_spreadsheet, _sheet_name, _cell_address, _angle), do: error()

  @spec set_cell_indent(reference(), String.t(), String.t(), integer()) :: :ok | {:error, atom()}
  def set_cell_indent(_spreadsheet, _sheet_name, _cell_address, _indent), do: error()

  @spec set_cell_alignment(reference(), String.t(), String.t(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_cell_alignment(_spreadsheet, _sheet_name, _cell_address, _horizontal, _vertical), do: error()

  # Image operations
  @spec add_image(reference(), String.t(), String.t(), String.t()) :: :ok | {:error, atom()}
  def add_image(_spreadsheet, _sheet_name, _cell_address, _image_path), do: error()

  @spec download_image(reference(), String.t(), String.t(), String.t()) :: :ok | {:error, atom()}
  def download_image(_spreadsheet, _sheet_name, _cell_address, _output_path), do: error()

  @spec change_image(reference(), String.t(), String.t(), String.t()) :: :ok | {:error, atom()}
  def change_image(_spreadsheet, _sheet_name, _cell_address, _new_image_path), do: error()

  # Chart operations
  @spec add_chart(
          reference(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          [String.t()],
          [String.t()],
          [String.t()]
        ) :: :ok | {:error, atom()}
  def add_chart(
        _spreadsheet,
        _sheet_name,
        _chart_type,
        _from_cell,
        _to_cell,
        _title,
        _data_series,
        _series_titles,
        _point_titles
      ),
      do: error()

  @spec add_chart_with_options(
          reference(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          [String.t()],
          [String.t()],
          [String.t()],
          integer(),
          boolean(),
          term(),
          term(),
          term(),
          term(),
          term()
        ) :: :ok | {:error, atom()}
  def add_chart_with_options(
        _spreadsheet,
        _sheet_name,
        _chart_type,
        _from_cell,
        _to_cell,
        _title,
        _data_series,
        _series_titles,
        _point_titles,
        _style,
        _vary_colors,
        _view_3d,
        _legend,
        _axes,
        _data_labels,
        _chart_specific
      ),
      do: error()

  @spec set_chart_style(reference(), String.t(), integer(), integer()) :: :ok | {:error, atom()}
  def set_chart_style(_spreadsheet, _sheet_name, _chart_index, _style), do: error()

  @spec set_chart_data_labels(
          reference(),
          String.t(),
          integer(),
          boolean(),
          boolean(),
          boolean(),
          boolean(),
          String.t()
        ) :: :ok | {:error, atom()}
  def set_chart_data_labels(
        _spreadsheet,
        _sheet_name,
        _chart_index,
        _show_values,
        _show_percent,
        _show_category_name,
        _show_series_name,
        _position
      ),
      do: error()

  @spec set_chart_legend_position(reference(), String.t(), integer(), String.t(), boolean()) ::
          :ok | {:error, atom()}
  def set_chart_legend_position(_spreadsheet, _sheet_name, _chart_index, _position, _overlay),
    do: error()

  @spec set_chart_3d_view(
          reference(),
          String.t(),
          integer(),
          integer(),
          integer(),
          integer(),
          integer()
        ) :: :ok | {:error, atom()}
  def set_chart_3d_view(
        _spreadsheet,
        _sheet_name,
        _chart_index,
        _rot_x,
        _rot_y,
        _perspective,
        _height_percent
      ),
      do: error()

  @spec set_chart_axis_titles(reference(), String.t(), integer(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def set_chart_axis_titles(
        _spreadsheet,
        _sheet_name,
        _chart_index,
        _category_axis_title,
        _value_axis_title
      ),
      do: error()

  # Cell/row/column operations
  @spec copy_row_styling(
          reference(),
          String.t(),
          integer(),
          integer(),
          integer() | nil,
          integer() | nil
        ) :: :ok | {:error, atom()}
  def copy_row_styling(
        _spreadsheet,
        _sheet_name,
        _source_row,
        _target_row,
        _start_column,
        _end_column
      ),
      do: error()

  @spec copy_column_styling(
          reference(),
          String.t(),
          integer(),
          integer(),
          integer() | nil,
          integer() | nil
        ) :: :ok | {:error, atom()}
  def copy_column_styling(
        _spreadsheet,
        _sheet_name,
        _source_column,
        _target_column,
        _start_row,
        _end_row
      ),
      do: error()

  @spec set_wrap_text(reference(), String.t(), String.t(), boolean()) :: :ok | {:error, atom()}
  def set_wrap_text(_spreadsheet, _sheet_name, _cell_address, _wrap), do: error()

  @spec set_column_width(reference(), String.t(), String.t(), float()) :: :ok | {:error, atom()}
  def set_column_width(_spreadsheet, _sheet_name, _column, _width), do: error()

  @spec set_column_auto_width(reference(), String.t(), String.t(), boolean()) ::
          :ok | {:error, atom()}
  def set_column_auto_width(_spreadsheet, _sheet_name, _column, _auto_width), do: error()

  # Sheet & workbook protection
  @spec set_sheet_protection(reference(), String.t(), String.t() | nil, boolean()) ::
          :ok | {:error, atom()}
  def set_sheet_protection(_spreadsheet, _sheet_name, _password, _is_protected), do: error()

  @spec set_workbook_protection(reference(), String.t()) :: :ok | {:error, atom()}
  def set_workbook_protection(_spreadsheet, _password), do: error()

  @spec set_sheet_state(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_sheet_state(_spreadsheet, _sheet_name, _state), do: error()

  @spec add_merge_cells(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def add_merge_cells(_spreadsheet, _sheet_name, _range), do: error()

  @spec set_password(String.t(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_password(_input_path, _output_path, _password), do: error()

  # Conditional formatting operations
  @spec add_cell_value_rule(
          reference(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          String.t() | nil,
          String.t()
        ) :: :ok | {:error, atom()}
  def add_cell_value_rule(
        _spreadsheet,
        _sheet_name,
        _cell_range,
        _operator,
        _value1,
        _value2,
        _format_style
      ),
      do: error()

  @spec add_color_scale(
          reference(),
          String.t(),
          String.t(),
          String.t(),
          String.t() | nil,
          map() | nil,
          String.t() | nil,
          String.t() | nil,
          map() | nil,
          String.t(),
          String.t() | nil,
          map() | nil
        ) :: :ok | {:error, atom()} | boolean()
  def add_color_scale(
    _spreadsheet,
    _sheet_name,
    _range,
    _min_type,
    _min_value,
    _min_color,
    _mid_type,
    _mid_value,
    _mid_color,
    _max_type,
    _max_value,
    _max_color
),
      do: error()

  @spec add_data_bar(
          reference(),
          String.t(),
          String.t(),
          {String.t(), String.t()} | nil,
          {String.t(), String.t()} | nil,
          String.t()
        ) :: :ok | {:error, atom()}
  def add_data_bar(_spreadsheet, _sheet_name, _cell_range, _min_value, _max_value, _color),
    do: error()

  @spec add_top_bottom_rule(reference(), String.t(), String.t(), String.t(), integer(), boolean(), String.t()) ::
          :ok | {:error, atom()}
  def add_top_bottom_rule(
        _spreadsheet,
        _sheet_name,
        _cell_range,
        _rule_type,
        _rank,
        _percent,
        _format_style
      ),
      do: error()

  @spec add_cell_is_rule(reference(), String.t(), String.t(), String.t(), String.t(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def add_cell_is_rule(_spreadsheet, _sheet_name, _range, _operator, _value1, _value2, _format_style),
    do: error()

  @spec add_text_rule(reference(), String.t(), String.t(), String.t(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def add_text_rule(_spreadsheet, _sheet_name, _range, _operator, _text, _format_style),
    do: error()

  @spec add_icon_set(reference(), String.t(), String.t(), String.t(), [{String.t(), String.t()}]) ::
          :ok | {:error, atom()}
  def add_icon_set(_spreadsheet, _sheet_name, _range, _icon_style, _thresholds),
    do: error()

  @spec add_above_below_average_rule(reference(), String.t(), String.t(), String.t(), integer() | nil, String.t()) ::
          :ok | {:error, atom()}
  def add_above_below_average_rule(_spreadsheet, _sheet_name, _range, _rule_type, _std_dev, _format_style),
    do: error()

  # Data validation operations
  @spec add_list_validation(
          reference(),
          String.t(),
          String.t(),
          [String.t()],
          boolean(),
          String.t(),
          String.t(),
          String.t(),
          boolean()
        ) ::
          :ok | {:error, atom()}
  def add_list_validation(
        _spreadsheet,
        _sheet_name,
        _range,
        _options,
        _show_dropdown,
        _prompt_title,
        _prompt_text,
        _error_text,
        _show_error
      ),
      do: error()

  @spec add_list_validation(reference(), String.t(), String.t(), [String.t()]) ::
          :ok | {:error, atom()}
  def add_list_validation(_spreadsheet, _sheet_name, _range, _options), do: error()

  @spec add_number_validation(
          reference(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          boolean(),
          String.t(),
          String.t(),
          String.t(),
          boolean()
        ) ::
          :ok | {:error, atom()}
  def add_number_validation(
        _spreadsheet,
        _sheet_name,
        _range,
        _operator,
        _value1,
        _value2,
        _show_dropdown,
        _prompt_title,
        _prompt_text,
        _error_text,
        _show_error
      ),
      do: error()

  @spec add_number_validation(
          reference(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          String.t()
        ) ::
          :ok | {:error, atom()}
  def add_number_validation(
        _spreadsheet,
        _sheet_name,
        _range,
        _operator,
        _value1,
        _value2
      ),
      do: error()

  @spec add_date_validation(
          reference(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          boolean(),
          String.t(),
          String.t(),
          String.t(),
          boolean()
        ) ::
          :ok | {:error, atom()}
  def add_date_validation(
        _spreadsheet,
        _sheet_name,
        _range,
        _operator,
        _value1,
        _value2,
        _show_dropdown,
        _prompt_title,
        _prompt_text,
        _error_text,
        _show_error
      ),
      do: error()

  @spec add_date_validation(
          reference(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          String.t()
        ) ::
          :ok | {:error, atom()}
  def add_date_validation(
        _spreadsheet,
        _sheet_name,
        _range,
        _operator,
        _value1,
        _value2
      ),
      do: error()

  @spec add_text_length_validation(
          reference(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          boolean(),
          String.t(),
          String.t(),
          String.t(),
          boolean()
        ) ::
          :ok | {:error, atom()}
  def add_text_length_validation(
        _spreadsheet,
        _sheet_name,
        _range,
        _operator,
        _value1,
        _value2,
        _show_dropdown,
        _prompt_title,
        _prompt_text,
        _error_text,
        _show_error
      ),
      do: error()

  @spec add_text_length_validation(
          reference(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          String.t()
        ) ::
          :ok | {:error, atom()}
  def add_text_length_validation(
        _spreadsheet,
        _sheet_name,
        _range,
        _operator,
        _value1,
        _value2
      ),
      do: error()

  @spec add_custom_validation(
          reference(),
          String.t(),
          String.t(),
          String.t(),
          boolean(),
          String.t(),
          String.t(),
          String.t(),
          boolean()
        ) ::
          :ok | {:error, atom()}
  def add_custom_validation(
        _spreadsheet,
        _sheet_name,
        _range,
        _formula,
        _show_dropdown,
        _prompt_title,
        _prompt_text,
        _error_text,
        _show_error
      ),
      do: error()

  @spec add_custom_validation(
          reference(),
          String.t(),
          String.t(),
          String.t()
        ) ::
          :ok | {:error, atom()}
  def add_custom_validation(
        _spreadsheet,
        _sheet_name,
        _range,
        _formula
      ),
      do: error()

  @spec remove_data_validation(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def remove_data_validation(_spreadsheet, _sheet_name, _range), do: error()

  # Pivot table operations
  @spec add_pivot_table(
          reference(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          [integer()],
          [integer()],
          [{integer(), String.t(), String.t()}]
        ) :: :ok | {:error, atom()}
  def add_pivot_table(
        _spreadsheet,
        _sheet_name,
        _name,
        _source_sheet,
        _source_range,
        _target_cell,
        _row_fields,
        _column_fields,
        _data_fields
      ),
      do: error()

  @spec has_pivot_tables(reference(), String.t()) :: boolean()
  def has_pivot_tables(_spreadsheet, _sheet_name), do: error()

  @spec count_pivot_tables(reference(), String.t()) :: {:ok, integer()} | {:error, atom()}
  def count_pivot_tables(_spreadsheet, _sheet_name), do: error()

  @spec refresh_all_pivot_tables(reference()) :: :ok | {:error, atom()}
  def refresh_all_pivot_tables(_spreadsheet), do: error()

  @spec remove_pivot_table(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def remove_pivot_table(_spreadsheet, _sheet_name, _pivot_table_name), do: error()

  # Table operations
  @spec add_table(
          reference(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          [String.t()],
          boolean() | nil
        ) :: :ok | {:error, atom()}
  def add_table(
        _spreadsheet,
        _sheet_name,
        _table_name,
        _display_name,
        _start_cell,
        _end_cell,
        _columns,
        _has_totals_row
      ),
      do: error()

  @spec get_tables(reference(), String.t()) :: [map()] | {:error, atom()}
  def get_tables(_spreadsheet, _sheet_name), do: error()

  @spec remove_table(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def remove_table(_spreadsheet, _sheet_name, _table_name), do: error()

  @spec has_tables(reference(), String.t()) :: boolean() | {:error, atom()}
  def has_tables(_spreadsheet, _sheet_name), do: error()

  @spec count_tables(reference(), String.t()) :: non_neg_integer() | {:error, atom()}
  def count_tables(_spreadsheet, _sheet_name), do: error()

  @spec set_table_style(
          reference(),
          String.t(),
          String.t(),
          String.t(),
          boolean(),
          boolean(),
          boolean(),
          boolean()
        ) :: :ok | {:error, atom()}
  def set_table_style(
        _spreadsheet,
        _sheet_name,
        _table_name,
        _style_name,
        _show_first_col,
        _show_last_col,
        _show_row_stripes,
        _show_col_stripes
      ),
      do: error()

  @spec remove_table_style(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def remove_table_style(_spreadsheet, _sheet_name, _table_name), do: error()

  @spec add_table_column(
          reference(),
          String.t(),
          String.t(),
          String.t(),
          String.t() | nil,
          String.t() | nil
        ) :: :ok | {:error, atom()}
  def add_table_column(
        _spreadsheet,
        _sheet_name,
        _table_name,
        _column_name,
        _totals_row_function,
        _totals_row_label
      ),
      do: error()

  @spec modify_table_column(
          reference(),
          String.t(),
          String.t(),
          String.t(),
          String.t() | nil,
          String.t() | nil,
          String.t() | nil
        ) :: :ok | {:error, atom()}
  def modify_table_column(
        _spreadsheet,
        _sheet_name,
        _table_name,
        _old_column_name,
        _new_column_name,
        _totals_row_function,
        _totals_row_label
      ),
      do: error()

  @spec set_table_totals_row(reference(), String.t(), String.t(), boolean()) ::
          :ok | {:error, atom()}
  def set_table_totals_row(_spreadsheet, _sheet_name, _table_name, _show_totals_row), do: error()

  # Drawing operations
  @spec add_shape(
          reference(),
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
        _spreadsheet,
        _sheet_name,
        _cell_address,
        _shape_type,
        _width,
        _height,
        _fill_color,
        _outline_color,
        _outline_width
      ),
      do: error()

  @spec add_text_box(
          reference(),
          String.t(),
          String.t(),
          String.t(),
          integer(),
          integer(),
          String.t(),
          String.t(),
          String.t(),
          float()
        ) :: :ok | {:error, atom()}
  def add_text_box(
        _spreadsheet,
        _sheet_name,
        _cell,
        _text,
        _width,
        _height,
        _fill_color,
        _text_color,
        _border_color,
        _border_width
      ),
      do: error()

  @spec add_connector(reference(), String.t(), String.t(), String.t(), String.t(), float()) ::
          :ok | {:error, atom()}
  def add_connector(_spreadsheet, _sheet_name, _from_cell, _to_cell, _color, _width), do: error()

  # Print settings functions
  @spec set_page_orientation(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_page_orientation(_spreadsheet, _sheet_name, _orientation), do: error()

  @spec set_paper_size(reference(), String.t(), integer()) :: :ok | {:error, atom()}
  def set_paper_size(_spreadsheet, _sheet_name, _paper_size), do: error()

  @spec set_page_scale(reference(), String.t(), integer()) :: :ok | {:error, atom()}
  def set_page_scale(_spreadsheet, _sheet_name, _scale), do: error()

  @spec set_fit_to_page(reference(), String.t(), integer(), integer()) :: :ok | {:error, atom()}
  def set_fit_to_page(_spreadsheet, _sheet_name, _width, _height), do: error()

  @spec set_page_margins(reference(), String.t(), float(), float(), float(), float()) :: :ok | {:error, atom()}
  def set_page_margins(_spreadsheet, _sheet_name, _top, _right, _bottom, _left), do: error()

  @spec set_header_footer_margins(reference(), String.t(), float(), float()) :: :ok | {:error, atom()}
  def set_header_footer_margins(_spreadsheet, _sheet_name, _header, _footer), do: error()

  @spec set_header(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_header(_spreadsheet, _sheet_name, _header), do: error()

  @spec set_footer(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_footer(_spreadsheet, _sheet_name, _footer), do: error()

  @spec set_print_centered(reference(), String.t(), boolean(), boolean()) :: :ok | {:error, atom()}
  def set_print_centered(_spreadsheet, _sheet_name, _horizontal, _vertical), do: error()

  @spec set_print_area(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_print_area(_spreadsheet, _sheet_name, _print_area), do: error()

  @spec set_print_titles(reference(), String.t(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_print_titles(_spreadsheet, _sheet_name, _rows, _columns), do: error()

  # Sheet view functions
  @spec set_show_grid_lines(reference(), String.t(), boolean()) :: :ok | {:error, atom()}
  def set_show_grid_lines(_spreadsheet, _sheet_name, _value), do: error()

  @spec set_tab_selected(reference(), String.t(), boolean()) :: :ok | {:error, atom()}
  def set_tab_selected(_spreadsheet, _sheet_name, _value), do: error()

  @spec set_top_left_cell(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_top_left_cell(_spreadsheet, _sheet_name, _cell_address), do: error()

  @spec set_zoom_scale(reference(), String.t(), integer()) :: :ok | {:error, atom()}
  def set_zoom_scale(_spreadsheet, _sheet_name, _value), do: error()

  @doc """
  Sets the view type for a worksheet (normal, page layout, page break preview).
  """
  @spec set_sheet_view(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_sheet_view(_spreadsheet, _sheet_name, _view_type), do: error()

  @doc """
  Sets zoom scale for normal view.
  """
  @spec set_zoom_scale_normal(reference(), String.t(), integer()) :: :ok | {:error, atom()}
  def set_zoom_scale_normal(_spreadsheet, _sheet_name, _scale), do: error()

  @doc """
  Sets zoom scale for page layout view.
  """
  @spec set_zoom_scale_page_layout(reference(), String.t(), integer()) :: :ok | {:error, atom()}
  def set_zoom_scale_page_layout(_spreadsheet, _sheet_name, _scale), do: error()

  @doc """
  Sets zoom scale for page break preview.
  """
  @spec set_zoom_scale_page_break(reference(), String.t(), integer()) :: :ok | {:error, atom()}
  def set_zoom_scale_page_break(_spreadsheet, _sheet_name, _scale), do: error()

  @doc """
  Split panes at the specified position.
  """
  @spec split_panes(reference(), String.t(), float(), float()) :: :ok | {:error, atom()}
  def split_panes(_spreadsheet, _sheet_name, _horizontal_position, _vertical_position), do: error()

  @doc """
  Freeze panes at the specified rows and columns.
  """
  @spec freeze_panes(reference(), String.t(), integer(), integer()) :: :ok | {:error, atom()}
  def freeze_panes(_spreadsheet, _sheet_name, _rows, _cols), do: error()

  @doc """
  Set the tab color for a sheet.
  """
  @spec set_tab_color(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_tab_color(_spreadsheet, _sheet_name, _color), do: error()

  @doc """
  Sets the active selection in a worksheet.
  """
  @spec set_selection(reference(), String.t(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_selection(_spreadsheet, _sheet_name, _active_cell, _sqref), do: error()

  # Workbook view functions
  @doc """
  Sets the active tab (sheet) when opening the workbook.
  """
  @spec set_active_tab(reference(), integer()) :: :ok | {:error, atom()}
  def set_active_tab(_spreadsheet, _tab_index), do: error()

  @doc """
  Sets the window position and size for the workbook.
  """
  @spec set_workbook_window_position(reference(), integer(), integer(), integer(), integer()) :: :ok | {:error, atom()}
  def set_workbook_window_position(_spreadsheet, _x_position, _y_position, _window_width, _window_height), do: error()

  # Comment functions
  @doc """
  Adds a comment to a cell.
  """
  @spec add_comment(reference(), String.t(), String.t(), String.t(), String.t()) :: :ok | {:error, atom()}
  def add_comment(_spreadsheet, _sheet_name, _cell_address, _text, _author), do: error()

  @doc """
  Gets the comment from a cell.
  """
  @spec get_comment(reference(), String.t(), String.t()) :: {:ok, String.t(), String.t()} | {:error, atom()}
  def get_comment(_spreadsheet, _sheet_name, _cell_address), do: error()

  @doc """
  Updates an existing comment in a cell.
  """
  @spec update_comment(reference(), String.t(), String.t(), String.t(), String.t() | nil) :: :ok | {:error, atom()}
  def update_comment(_spreadsheet, _sheet_name, _cell_address, _text, _author \\ nil), do: error()

  @doc """
  Removes a comment from a cell.
  """
  @spec remove_comment(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def remove_comment(_spreadsheet, _sheet_name, _cell_address), do: error()

  @doc """
  Checks if a sheet has any comments.
  """
  @spec has_comments(reference(), String.t()) :: boolean() | {:error, atom()}
  def has_comments(_spreadsheet, _sheet_name), do: error()

  @doc """
  Gets the number of comments in a sheet.
  """
  @spec get_comments_count(reference(), String.t()) :: integer() | {:error, atom()}
  def get_comments_count(_spreadsheet, _sheet_name), do: error()

  # Formula functions

  @doc """
  Sets a regular formula in a cell.
  """
  @spec set_formula(reference(), String.t(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_formula(_spreadsheet, _sheet_name, _cell_address, _formula), do: error()

  @doc """
  Sets an array formula for a range of cells.
  """
  @spec set_array_formula(reference(), String.t(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_array_formula(_spreadsheet, _sheet_name, _range, _formula), do: error()

  @doc """
  Creates a named range in the spreadsheet.
  """
  @spec create_named_range(reference(), String.t(), String.t(), String.t()) :: :ok | {:error, atom()}
  def create_named_range(_spreadsheet, _name, _sheet_name, _range), do: error()

  @doc """
  Creates a defined name in the spreadsheet.
  """
  @spec create_defined_name(reference(), String.t(), String.t(), String.t() | nil) :: :ok | {:error, atom()}
  def create_defined_name(_spreadsheet, _name, _formula, _sheet_name \\ nil), do: error()

  @doc """
  Gets all defined names in the spreadsheet.
  """
  @spec get_defined_names(reference()) :: [{String.t(), String.t()}] | {:error, atom()}
  def get_defined_names(_spreadsheet), do: error()

  @doc """
  Sets an auto filter for a range of cells in a worksheet.
  """
  @spec set_auto_filter(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_auto_filter(_spreadsheet, _sheet_name, _range), do: error()

  @doc """
  Removes an auto filter from a worksheet.
  """
  @spec remove_auto_filter(reference(), String.t()) :: :ok | {:error, atom()}
  def remove_auto_filter(_spreadsheet, _sheet_name), do: error()

  @doc """
  Checks if a worksheet has an auto filter.
  """
  @spec has_auto_filter(reference(), String.t()) :: boolean() | {:error, atom()}
  def has_auto_filter(_spreadsheet, _sheet_name), do: error()

  @doc """
  Gets the range of an auto filter in a worksheet.
  """
  @spec get_auto_filter_range(reference(), String.t()) :: String.t() | nil | {:error, atom()}
  def get_auto_filter_range(_spreadsheet, _sheet_name), do: error()

  defp error(), do: :erlang.nif_error(:nif_not_loaded)
end
