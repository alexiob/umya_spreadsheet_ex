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
  @spec write_with_compression(reference(), String.t(), non_neg_integer()) ::
          :ok | {:error, atom()}
  def write_with_compression(_spreadsheet, _path, _compression_level), do: error()

  @spec write_with_encryption_options(
          reference(),
          String.t(),
          String.t(),
          String.t(),
          String.t() | nil,
          non_neg_integer() | nil
        ) :: :ok | {:error, atom()}
  def write_with_encryption_options(
        _spreadsheet,
        _path,
        _password,
        _algorithm,
        _salt_value \\ nil,
        _spin_count \\ nil
      ),
      do: error()

  @spec to_binary_xlsx(reference()) :: binary() | {:error, atom()}
  def to_binary_xlsx(_spreadsheet), do: error()

  # CSV export functions
  @spec write_csv(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def write_csv(_spreadsheet, _sheet_name, _path), do: error()

  @spec write_csv_with_options(
          reference(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          boolean(),
          String.t()
        ) :: :ok | {:error, atom()}
  def write_csv_with_options(
        _spreadsheet,
        _sheet_name,
        _path,
        _encoding,
        _delimiter,
        _do_trim,
        _wrap_with_char
      ),
      do: error()

  # Aliased functions for compatibility between Rust and Elixir naming
  @spec write_light(reference(), String.t()) :: :ok | {:error, atom()}
  def write_light(spreadsheet, path), do: write_file_light(spreadsheet, path)

  @spec write_with_password_light(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def write_with_password_light(spreadsheet, path, password),
    do: write_file_with_password_light(spreadsheet, path, password)

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

  @spec set_font_underline(reference(), String.t(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def set_font_underline(_spreadsheet, _sheet_name, _cell_address, _underline_style), do: error()

  @spec set_font_strikethrough(reference(), String.t(), String.t(), boolean()) ::
          :ok | {:error, atom()}
  def set_font_strikethrough(_spreadsheet, _sheet_name, _cell_address, _is_strikethrough),
    do: error()

  @spec set_font_family(reference(), String.t(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def set_font_family(_spreadsheet, _sheet_name, _cell_address, _font_family),
    do: error()

  @spec set_font_scheme(reference(), String.t(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def set_font_scheme(_spreadsheet, _sheet_name, _cell_address, _font_scheme),
    do: error()

  @spec set_border_style(reference(), String.t(), String.t(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def set_border_style(_spreadsheet, _sheet_name, _cell_address, _border_position, _border_style),
    do: error()

  @spec set_cell_rotation(reference(), String.t(), String.t(), integer()) ::
          :ok | {:error, atom()}
  def set_cell_rotation(_spreadsheet, _sheet_name, _cell_address, _angle), do: error()

  @spec set_cell_indent(reference(), String.t(), String.t(), integer()) :: :ok | {:error, atom()}
  def set_cell_indent(_spreadsheet, _sheet_name, _cell_address, _indent), do: error()

  # Font formatting getter functions
  @spec get_font_name(reference(), String.t(), String.t()) :: String.t() | {:error, atom()}
  def get_font_name(_spreadsheet, _sheet_name, _cell_address), do: error()

  @spec get_font_size(reference(), String.t(), String.t()) :: number() | {:error, atom()}
  def get_font_size(_spreadsheet, _sheet_name, _cell_address), do: error()

  @spec get_font_bold(reference(), String.t(), String.t()) :: boolean() | {:error, atom()}
  def get_font_bold(_spreadsheet, _sheet_name, _cell_address), do: error()

  @spec get_font_italic(reference(), String.t(), String.t()) :: boolean() | {:error, atom()}
  def get_font_italic(_spreadsheet, _sheet_name, _cell_address), do: error()

  @spec get_font_underline(reference(), String.t(), String.t()) :: String.t() | {:error, atom()}
  def get_font_underline(_spreadsheet, _sheet_name, _cell_address), do: error()

  @spec get_font_strikethrough(reference(), String.t(), String.t()) ::
          boolean() | {:error, atom()}
  def get_font_strikethrough(_spreadsheet, _sheet_name, _cell_address), do: error()

  @spec get_font_family(reference(), String.t(), String.t()) :: String.t() | {:error, atom()}
  def get_font_family(_spreadsheet, _sheet_name, _cell_address), do: error()

  @spec get_font_scheme(reference(), String.t(), String.t()) :: String.t() | {:error, atom()}
  def get_font_scheme(_spreadsheet, _sheet_name, _cell_address), do: error()

  @spec get_font_color(reference(), String.t(), String.t()) :: String.t() | {:error, atom()}
  def get_font_color(_spreadsheet, _sheet_name, _cell_address), do: error()

  # Alignment getter functions
  @spec get_cell_horizontal_alignment(reference(), String.t(), String.t()) ::
          String.t() | {:error, atom()}
  def get_cell_horizontal_alignment(_spreadsheet, _sheet_name, _cell_address), do: error()

  @spec get_cell_vertical_alignment(reference(), String.t(), String.t()) ::
          String.t() | {:error, atom()}
  def get_cell_vertical_alignment(_spreadsheet, _sheet_name, _cell_address), do: error()

  @spec get_cell_wrap_text(reference(), String.t(), String.t()) :: boolean() | {:error, atom()}
  def get_cell_wrap_text(_spreadsheet, _sheet_name, _cell_address), do: error()

  @spec get_cell_text_rotation(reference(), String.t(), String.t()) ::
          non_neg_integer() | {:error, atom()}
  def get_cell_text_rotation(_spreadsheet, _sheet_name, _cell_address), do: error()

  # Border getter functions
  @spec get_border_style(reference(), String.t(), String.t(), String.t()) ::
          String.t() | {:error, atom()}
  def get_border_style(_spreadsheet, _sheet_name, _cell_address, _border_position), do: error()

  @spec get_border_color(reference(), String.t(), String.t(), String.t()) ::
          String.t() | {:error, atom()}
  def get_border_color(_spreadsheet, _sheet_name, _cell_address, _border_position), do: error()

  # Fill/background getter functions
  @spec get_cell_background_color(reference(), String.t(), String.t()) ::
          String.t() | {:error, atom()}
  def get_cell_background_color(_spreadsheet, _sheet_name, _cell_address), do: error()

  @spec get_cell_foreground_color(reference(), String.t(), String.t()) ::
          String.t() | {:error, atom()}
  def get_cell_foreground_color(_spreadsheet, _sheet_name, _cell_address), do: error()

  @spec get_cell_pattern_type(reference(), String.t(), String.t()) ::
          String.t() | {:error, atom()}
  def get_cell_pattern_type(_spreadsheet, _sheet_name, _cell_address), do: error()

  # Number format getter functions
  @spec get_cell_number_format_id(reference(), String.t(), String.t()) ::
          non_neg_integer() | {:error, atom()}
  def get_cell_number_format_id(_spreadsheet, _sheet_name, _cell_address), do: error()

  @spec get_cell_format_code(reference(), String.t(), String.t()) :: String.t() | {:error, atom()}
  def get_cell_format_code(_spreadsheet, _sheet_name, _cell_address), do: error()

  # Protection getter functions
  @spec get_cell_locked(reference(), String.t(), String.t()) :: boolean() | {:error, atom()}
  def get_cell_locked(_spreadsheet, _sheet_name, _cell_address), do: error()

  @spec get_cell_hidden(reference(), String.t(), String.t()) :: boolean() | {:error, atom()}
  def get_cell_hidden(_spreadsheet, _sheet_name, _cell_address), do: error()

  @spec set_cell_alignment(reference(), String.t(), String.t(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def set_cell_alignment(_spreadsheet, _sheet_name, _cell_address, _horizontal, _vertical),
    do: error()

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
          String.t() | nil,
          map() | nil,
          String.t() | nil,
          String.t() | nil,
          map() | nil,
          String.t() | nil,
          String.t() | nil,
          map() | nil,
          String.t() | nil
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

  @spec add_top_bottom_rule(
          reference(),
          String.t(),
          String.t(),
          String.t(),
          integer(),
          boolean(),
          String.t()
        ) ::
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

  @spec add_cell_is_rule(
          reference(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          String.t()
        ) ::
          :ok | {:error, atom()}
  def add_cell_is_rule(
        _spreadsheet,
        _sheet_name,
        _range,
        _operator,
        _value1,
        _value2,
        _format_style
      ),
      do: error()

  @spec add_text_rule(reference(), String.t(), String.t(), String.t(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def add_text_rule(_spreadsheet, _sheet_name, _range, _operator, _text, _format_style),
    do: error()

  @spec add_icon_set(reference(), String.t(), String.t(), String.t(), [{String.t(), String.t()}]) ::
          :ok | {:error, atom()}
  def add_icon_set(_spreadsheet, _sheet_name, _range, _icon_style, _thresholds),
    do: error()

  @spec add_above_below_average_rule(
          reference(),
          String.t(),
          String.t(),
          String.t(),
          integer() | nil,
          String.t()
        ) ::
          :ok | {:error, atom()}
  def add_above_below_average_rule(
        _spreadsheet,
        _sheet_name,
        _range,
        _rule_type,
        _std_dev,
        _format_style
      ),
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

  @spec get_table(reference(), String.t(), String.t()) :: {:ok, map()} | {:error, atom()}
  def get_table(_spreadsheet, _sheet_name, _table_name), do: error()

  @spec get_table_style(reference(), String.t(), String.t()) :: {:ok, map()} | {:error, atom()}
  def get_table_style(_spreadsheet, _sheet_name, _table_name), do: error()

  @spec get_table_columns(reference(), String.t(), String.t()) ::
          {:ok, [map()]} | {:error, atom()}
  def get_table_columns(_spreadsheet, _sheet_name, _table_name), do: error()

  @spec get_table_totals_row(reference(), String.t(), String.t()) ::
          {:ok, boolean()} | {:error, atom()}
  def get_table_totals_row(_spreadsheet, _sheet_name, _table_name), do: error()

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

  @spec set_page_margins(reference(), String.t(), float(), float(), float(), float()) ::
          :ok | {:error, atom()}
  def set_page_margins(_spreadsheet, _sheet_name, _top, _right, _bottom, _left), do: error()

  @spec set_header_footer_margins(reference(), String.t(), float(), float()) ::
          :ok | {:error, atom()}
  def set_header_footer_margins(_spreadsheet, _sheet_name, _header, _footer), do: error()

  @spec set_header(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_header(_spreadsheet, _sheet_name, _header), do: error()

  @spec set_footer(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_footer(_spreadsheet, _sheet_name, _footer), do: error()

  @spec set_print_centered(reference(), String.t(), boolean(), boolean()) ::
          :ok | {:error, atom()}
  def set_print_centered(_spreadsheet, _sheet_name, _horizontal, _vertical), do: error()

  @spec set_print_area(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_print_area(_spreadsheet, _sheet_name, _print_area), do: error()

  @spec set_print_titles(reference(), String.t(), String.t(), String.t()) ::
          :ok | {:error, atom()}
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
  def split_panes(_spreadsheet, _sheet_name, _horizontal_position, _vertical_position),
    do: error()

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
  Gets the active tab (sheet) index when opening the workbook.
  """
  @spec get_active_tab(reference()) :: {:ok, integer()} | {:error, atom()}
  def get_active_tab(_spreadsheet), do: error()

  @doc """
  Gets the window position and size for the workbook.
  """
  @spec get_workbook_window_position(reference()) :: {:ok, map()} | {:error, atom()}
  def get_workbook_window_position(_spreadsheet), do: error()

  @doc """
  Sets the active tab (sheet) when opening the workbook.
  """
  @spec set_active_tab(reference(), integer()) :: :ok | {:error, atom()}
  def set_active_tab(_spreadsheet, _tab_index), do: error()

  @doc """
  Sets the window position and size for the workbook.
  """
  @spec set_workbook_window_position(reference(), integer(), integer(), integer(), integer()) ::
          :ok | {:error, atom()}
  def set_workbook_window_position(
        _spreadsheet,
        _x_position,
        _y_position,
        _window_width,
        _window_height
      ),
      do: error()

  @doc """
  Checks if the workbook has protection enabled.
  """
  @spec is_workbook_protected(reference()) :: {:ok, boolean()} | {:error, atom()}
  def is_workbook_protected(_spreadsheet), do: error()

  @doc """
  Gets the workbook protection details including lock status.
  """
  @spec get_workbook_protection_details(reference()) :: {:ok, map()} | {:error, atom()}
  def get_workbook_protection_details(_spreadsheet), do: error()

  # Comment functions
  @doc """
  Adds a comment to a cell.
  """
  @spec add_comment(reference(), String.t(), String.t(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def add_comment(_spreadsheet, _sheet_name, _cell_address, _text, _author), do: error()

  @doc """
  Gets the comment from a cell.
  """
  @spec get_comment(reference(), String.t(), String.t()) ::
          {:ok, String.t(), String.t()} | {:error, atom()}
  def get_comment(_spreadsheet, _sheet_name, _cell_address), do: error()

  @doc """
  Updates an existing comment in a cell.
  """
  @spec update_comment(reference(), String.t(), String.t(), String.t(), String.t() | nil) ::
          :ok | {:error, atom()}
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

  # Hyperlink functions

  @doc """
  Adds a hyperlink to a cell.
  """
  @spec add_hyperlink(
          reference(),
          String.t(),
          String.t(),
          String.t(),
          String.t() | nil,
          boolean()
        ) :: :ok | {:error, atom()}
  def add_hyperlink(
        _spreadsheet,
        _sheet_name,
        _cell_address,
        _url,
        _tooltip \\ nil,
        _is_internal \\ false
      ),
      do: error()

  @doc """
  Gets hyperlink information from a cell.
  """
  @spec get_hyperlink(reference(), String.t(), String.t()) :: {:ok, map()} | {:error, atom()}
  def get_hyperlink(_spreadsheet, _sheet_name, _cell_address), do: error()

  @doc """
  Removes a hyperlink from a cell.
  """
  @spec remove_hyperlink(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def remove_hyperlink(_spreadsheet, _sheet_name, _cell_address), do: error()

  @doc """
  Checks if a specific cell has a hyperlink.
  """
  @spec has_hyperlink(reference(), String.t(), String.t()) :: boolean() | {:error, atom()}
  def has_hyperlink(_spreadsheet, _sheet_name, _cell_address), do: error()

  @doc """
  Checks if a worksheet contains any hyperlinks.
  """
  @spec has_hyperlinks(reference(), String.t()) :: boolean() | {:error, atom()}
  def has_hyperlinks(_spreadsheet, _sheet_name), do: error()

  @doc """
  Gets all hyperlinks from a worksheet.
  """
  @spec get_hyperlinks(reference(), String.t()) :: {:ok, list(map())} | {:error, atom()}
  def get_hyperlinks(_spreadsheet, _sheet_name), do: error()

  @doc """
  Updates an existing hyperlink in a cell.
  """
  @spec update_hyperlink(
          reference(),
          String.t(),
          String.t(),
          String.t(),
          String.t() | nil,
          boolean()
        ) :: :ok | {:error, atom()}
  def update_hyperlink(
        _spreadsheet,
        _sheet_name,
        _cell_address,
        _url,
        _tooltip \\ nil,
        _is_internal \\ false
      ),
      do: error()

  # Formula functions

  @doc """
  Sets a regular formula in a cell.
  """
  @spec set_formula(reference(), String.t(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_formula(_spreadsheet, _sheet_name, _cell_address, _formula), do: error()

  @doc """
  Sets an array formula for a range of cells.
  """
  @spec set_array_formula(reference(), String.t(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def set_array_formula(_spreadsheet, _sheet_name, _range, _formula), do: error()

  @doc """
  Creates a named range in the spreadsheet.
  """
  @spec create_named_range(reference(), String.t(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def create_named_range(_spreadsheet, _name, _sheet_name, _range), do: error()

  @doc """
  Creates a defined name in the spreadsheet.
  """
  @spec create_defined_name(reference(), String.t(), String.t(), String.t() | nil) ::
          :ok | {:error, atom()}
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

  # Rich Text functions

  @doc """
  Creates a new RichText object.
  """
  @spec create_rich_text() :: reference() | {:error, atom()}
  def create_rich_text(), do: error()

  @doc """
  Creates a RichText object from HTML string.
  """
  @spec create_rich_text_from_html(String.t()) :: reference() | {:error, atom()}
  def create_rich_text_from_html(_html), do: error()

  @doc """
  Creates a TextElement with text and optional font properties.
  """
  @spec create_text_element(String.t(), map()) :: reference() | {:error, atom()}
  def create_text_element(_text, _font_props), do: error()

  @doc """
  Gets text from a TextElement.
  """
  @spec get_text_element_text(reference()) :: String.t() | {:error, atom()}
  def get_text_element_text(_text_element), do: error()

  @doc """
  Gets font properties from a TextElement.
  """
  @spec get_text_element_font_properties(reference()) :: map() | {:error, atom()}
  def get_text_element_font_properties(_text_element), do: error()

  @doc """
  Adds a TextElement to a RichText object.
  """
  @spec add_text_element_to_rich_text(reference(), reference()) :: :ok | {:error, atom()}
  def add_text_element_to_rich_text(_rich_text, _text_element), do: error()

  @doc """
  Adds formatted text directly to a RichText object.
  """
  @spec add_formatted_text_to_rich_text(reference(), String.t(), map()) :: :ok | {:error, atom()}
  def add_formatted_text_to_rich_text(_rich_text, _text, _font_props), do: error()

  @doc """
  Sets rich text to a cell.
  """
  @spec set_cell_rich_text(reference(), String.t(), String.t(), reference()) ::
          :ok | {:error, atom()}
  def set_cell_rich_text(_spreadsheet, _sheet_name, _coordinate, _rich_text), do: error()

  @doc """
  Gets rich text from a cell.
  """
  @spec get_cell_rich_text(reference(), String.t(), String.t()) :: reference() | {:error, atom()}
  def get_cell_rich_text(_spreadsheet, _sheet_name, _coordinate), do: error()

  @doc """
  Gets plain text from RichText.
  """
  @spec get_rich_text_plain_text(reference()) :: String.t() | {:error, atom()}
  def get_rich_text_plain_text(_rich_text), do: error()

  @doc """
  Converts RichText to HTML.
  """
  @spec rich_text_to_html(reference()) :: String.t() | {:error, atom()}
  def rich_text_to_html(_rich_text), do: error()

  @doc """
  Gets text elements from RichText.
  """
  @spec get_rich_text_elements(reference()) :: [reference()] | {:error, atom()}
  def get_rich_text_elements(_rich_text), do: error()

  # OLE Objects functions

  @doc """
  Creates a new OLE objects collection.
  """
  @spec create_ole_objects() :: reference() | {:error, atom()}
  def create_ole_objects(), do: error()

  @doc """
  Creates a new OLE object.
  """
  @spec create_ole_object() :: reference() | {:error, atom()}
  def create_ole_object(), do: error()

  @doc """
  Creates new embedded object properties.
  """
  @spec create_embedded_object_properties() :: reference() | {:error, atom()}
  def create_embedded_object_properties(), do: error()

  @doc """
  Gets OLE objects from a worksheet.
  """
  @spec get_ole_objects(reference(), String.t()) :: reference() | {:error, atom()}
  def get_ole_objects(_spreadsheet, _sheet_name), do: error()

  @doc """
  Sets OLE objects to a worksheet.
  """
  @spec set_ole_objects(reference(), String.t(), reference()) :: :ok | {:error, atom()}
  def set_ole_objects(_spreadsheet, _sheet_name, _ole_objects), do: error()

  @doc """
  Adds an OLE object to a collection.
  """
  @spec add_ole_object(reference(), reference()) :: :ok | {:error, atom()}
  def add_ole_object(_ole_objects, _ole_object), do: error()

  @doc """
  Gets all OLE objects from a collection.
  """
  @spec get_ole_object_list(reference()) :: [reference()] | {:error, atom()}
  def get_ole_object_list(_ole_objects), do: error()

  @doc """
  Gets count of OLE objects in a collection.
  """
  @spec count_ole_objects(reference()) :: non_neg_integer() | {:error, atom()}
  def count_ole_objects(_ole_objects), do: error()

  @doc """
  Checks if collection has any OLE objects.
  """
  @spec has_ole_objects(reference()) :: boolean() | {:error, atom()}
  def has_ole_objects(_ole_objects), do: error()

  @doc """
  Gets requires property from OLE object.
  """
  @spec get_ole_object_requires(reference()) :: String.t() | {:error, atom()}
  def get_ole_object_requires(_ole_object), do: error()

  @doc """
  Sets requires property for OLE object.
  """
  @spec set_ole_object_requires(reference(), String.t()) :: :ok | {:error, atom()}
  def set_ole_object_requires(_ole_object, _requires), do: error()

  @doc """
  Gets program ID from OLE object.
  """
  @spec get_ole_object_prog_id(reference()) :: String.t() | {:error, atom()}
  def get_ole_object_prog_id(_ole_object), do: error()

  @doc """
  Sets program ID for OLE object.
  """
  @spec set_ole_object_prog_id(reference(), String.t()) :: :ok | {:error, atom()}
  def set_ole_object_prog_id(_ole_object, _prog_id), do: error()

  @doc """
  Gets object extension from OLE object.
  """
  @spec get_ole_object_extension(reference()) :: String.t() | {:error, atom()}
  def get_ole_object_extension(_ole_object), do: error()

  @doc """
  Sets object extension for OLE object.
  """
  @spec set_ole_object_extension(reference(), String.t()) :: :ok | {:error, atom()}
  def set_ole_object_extension(_ole_object, _extension), do: error()

  @doc """
  Gets object data from OLE object.
  """
  @spec get_ole_object_data(reference()) :: binary() | nil | {:error, atom()}
  def get_ole_object_data(_ole_object), do: error()

  @doc """
  Sets object data for OLE object.
  """
  @spec set_ole_object_data(reference(), binary()) :: :ok | {:error, atom()}
  def set_ole_object_data(_ole_object, _data), do: error()

  @doc """
  Gets embedded object properties from OLE object.
  """
  @spec get_ole_object_properties(reference()) :: reference() | {:error, atom()}
  def get_ole_object_properties(_ole_object), do: error()

  @doc """
  Sets embedded object properties for OLE object.
  """
  @spec set_ole_object_properties(reference(), reference()) :: :ok | {:error, atom()}
  def set_ole_object_properties(_ole_object, _properties), do: error()

  @doc """
  Gets program ID from embedded object properties.
  """
  @spec get_embedded_object_prog_id(reference()) :: String.t() | {:error, atom()}
  def get_embedded_object_prog_id(_properties), do: error()

  @doc """
  Sets program ID for embedded object properties.
  """
  @spec set_embedded_object_prog_id(reference(), String.t()) :: :ok | {:error, atom()}
  def set_embedded_object_prog_id(_properties, _prog_id), do: error()

  @doc """
  Gets shape ID from embedded object properties.
  """
  @spec get_embedded_object_shape_id(reference()) :: non_neg_integer() | {:error, atom()}
  def get_embedded_object_shape_id(_properties), do: error()

  @doc """
  Sets shape ID for embedded object properties.
  """
  @spec set_embedded_object_shape_id(reference(), non_neg_integer()) :: :ok | {:error, atom()}
  def set_embedded_object_shape_id(_properties, _shape_id), do: error()

  @doc """
  Loads OLE object from file.
  """
  @spec load_ole_object_from_file(String.t(), String.t()) :: reference() | {:error, atom()}
  def load_ole_object_from_file(_file_path, _prog_id), do: error()

  @doc """
  Saves OLE object data to file.
  """
  @spec save_ole_object_to_file(reference(), String.t()) :: :ok | {:error, atom()}
  def save_ole_object_to_file(_ole_object, _file_path), do: error()

  @doc """
  Creates OLE object with file data.
  """
  @spec create_ole_object_with_data(String.t(), String.t(), binary()) ::
          reference() | {:error, atom()}
  def create_ole_object_with_data(_prog_id, _extension, _data), do: error()

  @doc """
  Checks if OLE object is binary format.
  """
  @spec is_ole_object_binary(reference()) :: boolean() | {:error, atom()}
  def is_ole_object_binary(_ole_object), do: error()

  @doc """
  Checks if OLE object is Excel format.
  """
  @spec is_ole_object_excel(reference()) :: boolean() | {:error, atom()}
  def is_ole_object_excel(_ole_object), do: error()

  # Page Breaks functions
  @doc """
  Adds a row page break at the specified row number.
  """
  @spec add_row_page_break(reference(), String.t(), non_neg_integer(), boolean()) ::
          :ok | {:error, atom()}
  def add_row_page_break(_spreadsheet, _sheet_name, _row_number, _manual), do: error()

  @doc """
  Adds a column page break at the specified column number.
  """
  @spec add_column_page_break(reference(), String.t(), non_neg_integer(), boolean()) ::
          :ok | {:error, atom()}
  def add_column_page_break(_spreadsheet, _sheet_name, _column_number, _manual), do: error()

  @doc """
  Removes a row page break at the specified row number.
  """
  @spec remove_row_page_break(reference(), String.t(), non_neg_integer()) ::
          :ok | {:error, atom()}
  def remove_row_page_break(_spreadsheet, _sheet_name, _row_number), do: error()

  @doc """
  Removes a column page break at the specified column number.
  """
  @spec remove_column_page_break(reference(), String.t(), non_neg_integer()) ::
          :ok | {:error, atom()}
  def remove_column_page_break(_spreadsheet, _sheet_name, _column_number), do: error()

  @doc """
  Gets all row page breaks for the specified sheet.
  """
  @spec get_row_page_breaks(reference(), String.t()) ::
          [{non_neg_integer(), boolean()}] | {:error, atom()}
  def get_row_page_breaks(_spreadsheet, _sheet_name), do: error()

  @doc """
  Gets all column page breaks for the specified sheet.
  """
  @spec get_column_page_breaks(reference(), String.t()) ::
          [{non_neg_integer(), boolean()}] | {:error, atom()}
  def get_column_page_breaks(_spreadsheet, _sheet_name), do: error()

  @doc """
  Clears all row page breaks for the specified sheet.
  """
  @spec clear_row_page_breaks(reference(), String.t()) :: :ok | {:error, atom()}
  def clear_row_page_breaks(_spreadsheet, _sheet_name), do: error()

  @doc """
  Clears all column page breaks for the specified sheet.
  """
  @spec clear_column_page_breaks(reference(), String.t()) :: :ok | {:error, atom()}
  def clear_column_page_breaks(_spreadsheet, _sheet_name), do: error()

  @doc """
  Checks if a row page break exists at the specified row number.
  """
  @spec has_row_page_break(reference(), String.t(), non_neg_integer()) ::
          boolean() | {:error, atom()}
  def has_row_page_break(_spreadsheet, _sheet_name, _row_number), do: error()

  @doc """
  Checks if a column page break exists at the specified column number.
  """
  @spec has_column_page_break(reference(), String.t(), non_neg_integer()) ::
          boolean() | {:error, atom()}
  def has_column_page_break(_spreadsheet, _sheet_name, _column_number), do: error()

  # VML Support functions
  @doc """
  Creates a new VML shape in the specified worksheet.
  """
  @spec create_vml_shape(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def create_vml_shape(_spreadsheet, _sheet_name, _shape_id), do: error()

  @doc """
  Sets the CSS style for a VML shape.
  """
  @spec set_vml_shape_style(reference(), String.t(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def set_vml_shape_style(_spreadsheet, _sheet_name, _shape_id, _style), do: error()

  @doc """
  Sets the type of a VML shape.
  """
  @spec set_vml_shape_type(reference(), String.t(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def set_vml_shape_type(_spreadsheet, _sheet_name, _shape_id, _shape_type), do: error()

  @doc """
  Sets whether a VML shape is filled.
  """
  @spec set_vml_shape_filled(reference(), String.t(), String.t(), boolean()) ::
          :ok | {:error, atom()}
  def set_vml_shape_filled(_spreadsheet, _sheet_name, _shape_id, _filled), do: error()

  @doc """
  Sets the fill color for a VML shape.
  """
  @spec set_vml_shape_fill_color(reference(), String.t(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def set_vml_shape_fill_color(_spreadsheet, _sheet_name, _shape_id, _fill_color), do: error()

  @doc """
  Sets whether a VML shape has a stroke (outline).
  """
  @spec set_vml_shape_stroked(reference(), String.t(), String.t(), boolean()) ::
          :ok | {:error, atom()}
  def set_vml_shape_stroked(_spreadsheet, _sheet_name, _shape_id, _stroked), do: error()

  @doc """
  Sets the stroke color for a VML shape.
  """
  @spec set_vml_shape_stroke_color(reference(), String.t(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def set_vml_shape_stroke_color(_spreadsheet, _sheet_name, _shape_id, _stroke_color), do: error()

  @doc """
  Sets the stroke weight (thickness) for a VML shape.
  """
  @spec set_vml_shape_stroke_weight(reference(), String.t(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def set_vml_shape_stroke_weight(_spreadsheet, _sheet_name, _shape_id, _stroke_weight),
    do: error()

  # Document Properties functions

  @doc """
  Gets a custom document property.
  """
  @spec get_custom_property(reference(), String.t()) :: {:ok, String.t()} | {:error, atom()}
  def get_custom_property(_spreadsheet, _property_name), do: error()

  @doc """
  Sets a string custom document property.
  """
  @spec set_custom_property_string(reference(), String.t(), String.t()) :: :ok | {:error, atom()}
  def set_custom_property_string(_spreadsheet, _property_name, _value), do: error()

  @doc """
  Sets a number custom document property.
  """
  @spec set_custom_property_number(reference(), String.t(), integer()) :: :ok | {:error, atom()}
  def set_custom_property_number(_spreadsheet, _property_name, _value), do: error()

  @doc """
  Sets a boolean custom document property.
  """
  @spec set_custom_property_bool(reference(), String.t(), boolean()) :: :ok | {:error, atom()}
  def set_custom_property_bool(_spreadsheet, _property_name, _value), do: error()

  @doc """
  Sets a date custom document property.
  """
  @spec set_custom_property_date(reference(), String.t(), integer(), integer(), integer()) ::
          :ok | {:error, atom()}
  def set_custom_property_date(_spreadsheet, _property_name, _year, _month, _day), do: error()

  @doc """
  Removes a custom document property.
  """
  @spec remove_custom_property(reference(), String.t()) :: :ok | {:error, atom()}
  def remove_custom_property(_spreadsheet, _property_name), do: error()

  @doc """
  Gets all custom property names.
  """
  @spec get_custom_property_names(reference()) :: [String.t()] | {:error, atom()}
  def get_custom_property_names(_spreadsheet), do: error()

  @doc """
  Checks if a custom property exists.
  """
  @spec has_custom_property(reference(), String.t()) :: boolean() | {:error, atom()}
  def has_custom_property(_spreadsheet, _property_name), do: error()

  @doc """
  Gets the count of custom properties.
  """
  @spec get_custom_properties_count(reference()) :: integer() | {:error, atom()}
  def get_custom_properties_count(_spreadsheet), do: error()

  @doc """
  Clears all custom properties.
  """
  @spec clear_custom_properties(reference()) :: :ok | {:error, atom()}
  def clear_custom_properties(_spreadsheet), do: error()

  @doc """
  Gets the document title.
  """
  @spec get_title(reference()) :: {:ok, String.t()} | {:error, atom()}
  def get_title(_spreadsheet), do: error()

  @doc """
  Sets the document title.
  """
  @spec set_title(reference(), String.t()) :: :ok | {:error, atom()}
  def set_title(_spreadsheet, _title), do: error()

  @doc """
  Gets the document description.
  """
  @spec get_description(reference()) :: {:ok, String.t()} | {:error, atom()}
  def get_description(_spreadsheet), do: error()

  @doc """
  Sets the document description.
  """
  @spec set_description(reference(), String.t()) :: :ok | {:error, atom()}
  def set_description(_spreadsheet, _description), do: error()

  @doc """
  Gets the document subject.
  """
  @spec get_subject(reference()) :: {:ok, String.t()} | {:error, atom()}
  def get_subject(_spreadsheet), do: error()

  @doc """
  Sets the document subject.
  """
  @spec set_subject(reference(), String.t()) :: :ok | {:error, atom()}
  def set_subject(_spreadsheet, _subject), do: error()

  @doc """
  Gets the document keywords.
  """
  @spec get_keywords(reference()) :: {:ok, String.t()} | {:error, atom()}
  def get_keywords(_spreadsheet), do: error()

  @doc """
  Sets the document keywords.
  """
  @spec set_keywords(reference(), String.t()) :: :ok | {:error, atom()}
  def set_keywords(_spreadsheet, _keywords), do: error()

  @doc """
  Gets the document creator.
  """
  @spec get_creator(reference()) :: {:ok, String.t()} | {:error, atom()}
  def get_creator(_spreadsheet), do: error()

  @doc """
  Sets the document creator.
  """
  @spec set_creator(reference(), String.t()) :: :ok | {:error, atom()}
  def set_creator(_spreadsheet, _creator), do: error()

  @doc """
  Gets the document last modified by.
  """
  @spec get_last_modified_by(reference()) :: {:ok, String.t()} | {:error, atom()}
  def get_last_modified_by(_spreadsheet), do: error()

  @doc """
  Sets the document last modified by.
  """
  @spec set_last_modified_by(reference(), String.t()) :: :ok | {:error, atom()}
  def set_last_modified_by(_spreadsheet, _last_modified_by), do: error()

  @doc """
  Gets the document category.
  """
  @spec get_category(reference()) :: {:ok, String.t()} | {:error, atom()}
  def get_category(_spreadsheet), do: error()

  @doc """
  Sets the document category.
  """
  @spec set_category(reference(), String.t()) :: :ok | {:error, atom()}
  def set_category(_spreadsheet, _category), do: error()

  @doc """
  Gets the document company.
  """
  @spec get_company(reference()) :: {:ok, String.t()} | {:error, atom()}
  def get_company(_spreadsheet), do: error()

  @doc """
  Sets the document company.
  """
  @spec set_company(reference(), String.t()) :: :ok | {:error, atom()}
  def set_company(_spreadsheet, _company), do: error()

  @doc """
  Gets the document manager.
  """
  @spec get_manager(reference()) :: {:ok, String.t()} | {:error, atom()}
  def get_manager(_spreadsheet), do: error()

  @doc """
  Sets the document manager.
  """
  @spec set_manager(reference(), String.t()) :: :ok | {:error, atom()}
  def set_manager(_spreadsheet, _manager), do: error()

  @doc """
  Gets the document created date.
  """
  @spec get_created(reference()) :: {:ok, String.t()} | {:error, atom()}
  def get_created(_spreadsheet), do: error()

  @doc """
  Sets the document created date.
  """
  @spec set_created(reference(), String.t()) :: :ok | {:error, atom()}
  def set_created(_spreadsheet, _created), do: error()

  @doc """
  Gets the document modified date.
  """
  @spec get_modified(reference()) :: {:ok, String.t()} | {:error, atom()}
  def get_modified(_spreadsheet), do: error()

  @doc """
  Sets the document modified date.
  """
  @spec set_modified(reference(), String.t()) :: :ok | {:error, atom()}
  def set_modified(_spreadsheet, _modified), do: error()

  @spec test_simple_function() :: String.t()
  def test_simple_function(), do: error()

  defp error(), do: :erlang.nif_error(:nif_not_loaded)
end
