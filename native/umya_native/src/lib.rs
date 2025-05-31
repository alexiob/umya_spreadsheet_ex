use std::sync::Mutex;
use umya_spreadsheet;

// Import modules
mod auto_filter_functions;
mod cell_formatting;
mod cell_functions;
mod cell_operations;
mod chart_functions;
mod comment_functions;
pub mod conditional_formatting;
pub mod conditional_formatting_additional;
pub mod custom_structs;
mod data_validation;
mod document_properties;
mod drawing_functions;
mod file_format_options;
mod file_operations;
mod formula_functions;
mod get_cell_formatting;
mod helpers;
mod hyperlink;
mod image_functions;
mod ole_object_functions;
mod page_breaks;
mod password_functions;
mod pivot_table;
mod print_settings;
mod rich_text_functions;
mod row_column_operations;
mod set_background_color;
mod set_cell_alignment;
mod sheet_operations;
mod sheet_view_functions;
mod styling_operations;
mod table;
mod vml_support;
mod workbook_protection_functions;
mod workbook_view_functions;
mod write_csv_with_options;

// Define atoms for use in Elixir interface
mod atoms {
    rustler::atoms! {
        ok,
        error,
        not_found,
        invalid_path,
        unknown_error,

        // Chart function atoms
        position,
        overlay,
        right,
        left,
        top,
        bottom,
        top_right,
        rot_x,
        rot_y,
        perspective,
        height_percent,

        // Data label atoms
        show_values,
        show_percent,
        show_category_name,
        show_series_name,

        // Data validation atoms
        between,
        not_between,
        equal,
        not_equal,
        greater_than,
        less_than,
        greater_than_or_equal,
        less_than_or_equal,
    }
}

// Primary data structure representing a spreadsheet
pub struct UmyaSpreadsheet {
    spreadsheet: Mutex<umya_spreadsheet::Spreadsheet>,
}

// Register the NIF module
fn load(env: rustler::Env, _info: rustler::Term) -> bool {
    let _ = rustler::resource!(UmyaSpreadsheet, env);
    // Re-enabling Rich Text resources
    let _ = rustler::resource!(rich_text_functions::RichTextResource, env);
    let _ = rustler::resource!(rich_text_functions::TextElementResource, env);
    // OLE Objects resources
    let _ = rustler::resource!(ole_object_functions::OleObjectsResource, env);
    let _ = rustler::resource!(ole_object_functions::OleObjectResource, env);
    let _ = rustler::resource!(ole_object_functions::EmbeddedObjectPropertiesResource, env);
    true
}

rustler::init!(
    "Elixir.UmyaNative",
    load = load,
    nifs = [
        // File operations
        file_operations::new_file,
        file_operations::new_file_empty_worksheet,
        file_operations::read_file,
        file_operations::lazy_read_file,
        file_operations::write_file,
        file_operations::write_file_light,
        file_operations::write_file_with_password,
        file_operations::write_file_with_password_light,
        // Cell operations
        cell_operations::get_cell_value,
        cell_operations::set_cell_value,
        cell_operations::remove_cell,
        cell_operations::get_formatted_value,
        cell_operations::set_number_format,
        // Cell functions
        cell_functions::set_wrap_text,
        // Image functions
        image_functions::add_image,
        image_functions::download_image,
        image_functions::change_image,
        image_functions::get_image_dimensions,
        image_functions::list_images,
        image_functions::get_image_info,
        // Password functions
        password_functions::set_password,
        // Sheet operations
        sheet_operations::get_sheet_names,
        sheet_operations::get_sheet_count,
        sheet_operations::get_sheet_state,
        sheet_operations::get_sheet_protection,
        sheet_operations::get_merge_cells,
        sheet_operations::add_sheet,
        sheet_operations::clone_sheet,
        sheet_operations::remove_sheet,
        sheet_operations::rename_sheet,
        sheet_operations::insert_new_row,
        sheet_operations::insert_new_column,
        sheet_operations::set_sheet_protection,
        sheet_operations::add_merge_cells,
        sheet_operations::set_sheet_state,
        sheet_operations::move_range,
        // Styling operations
        styling_operations::set_font_color,
        styling_operations::set_font_size,
        styling_operations::set_font_bold,
        styling_operations::set_font_name,
        styling_operations::copy_row_styling,
        styling_operations::copy_column_styling,
        // Row/column operations
        row_column_operations::set_row_height,
        row_column_operations::set_row_style,
        row_column_operations::set_column_width,
        row_column_operations::set_column_auto_width,
        row_column_operations::get_column_width,
        row_column_operations::get_column_auto_width,
        row_column_operations::get_column_hidden,
        row_column_operations::get_row_height,
        row_column_operations::get_row_hidden,
        row_column_operations::remove_row,
        row_column_operations::remove_column,
        row_column_operations::remove_column_by_index,
        // Pivot table operations
        pivot_table::has_pivot_tables,
        pivot_table::add_pivot_table,
        pivot_table::count_pivot_tables,
        pivot_table::remove_pivot_table,
        pivot_table::refresh_all_pivot_tables,
        pivot_table::get_pivot_table_names,
        pivot_table::get_pivot_table_info,
        pivot_table::get_pivot_table_source_range,
        pivot_table::get_pivot_table_target_cell,
        pivot_table::get_pivot_table_fields,
        // Drawing and shape functions
        drawing_functions::add_shape,
        drawing_functions::add_text_box,
        drawing_functions::add_connector,
        // Cell formatting functions
        cell_formatting::set_font_italic,
        cell_formatting::set_font_underline,
        cell_formatting::set_font_strikethrough,
        cell_formatting::set_font_family,
        cell_formatting::set_font_scheme,
        cell_formatting::set_border_style,
        cell_formatting::set_cell_rotation,
        cell_formatting::set_cell_indent,
        // Cell operations functions
        cell_functions::set_wrap_text,
        // Image functions
        image_functions::add_image,
        image_functions::download_image,
        image_functions::change_image,
        image_functions::get_image_dimensions,
        image_functions::list_images,
        image_functions::get_image_info,
        // Cell formatting getter functions
        get_cell_formatting::get_font_name,
        get_cell_formatting::get_font_size,
        get_cell_formatting::get_font_bold,
        get_cell_formatting::get_font_italic,
        get_cell_formatting::get_font_underline,
        get_cell_formatting::get_font_strikethrough,
        get_cell_formatting::get_font_family,
        get_cell_formatting::get_font_scheme,
        get_cell_formatting::get_font_color,
        get_cell_formatting::get_cell_horizontal_alignment,
        get_cell_formatting::get_cell_vertical_alignment,
        get_cell_formatting::get_cell_wrap_text,
        get_cell_formatting::get_cell_text_rotation,
        get_cell_formatting::get_border_style,
        get_cell_formatting::get_border_color,
        get_cell_formatting::get_cell_background_color,
        get_cell_formatting::get_cell_foreground_color,
        get_cell_formatting::get_cell_pattern_type,
        get_cell_formatting::get_cell_number_format_id,
        get_cell_formatting::get_cell_format_code,
        get_cell_formatting::get_cell_locked,
        get_cell_formatting::get_cell_hidden,
        // Print settings functions
        print_settings::set_page_orientation,
        print_settings::set_paper_size,
        print_settings::set_page_scale,
        print_settings::set_fit_to_page,
        print_settings::set_page_margins,
        print_settings::set_header_footer_margins,
        print_settings::set_header,
        print_settings::set_footer,
        print_settings::set_print_centered,
        print_settings::set_print_area,
        print_settings::set_print_titles,
        print_settings::get_page_orientation,
        print_settings::get_paper_size,
        print_settings::get_page_scale,
        print_settings::get_fit_to_page,
        print_settings::get_page_margins,
        print_settings::get_header_footer_margins,
        print_settings::get_header,
        print_settings::get_footer,
        print_settings::get_print_centered,
        print_settings::get_print_area,
        print_settings::get_print_titles,
        // Conditional formatting functions
        conditional_formatting::add_cell_value_rule,
        conditional_formatting::add_color_scale,
        conditional_formatting_additional::add_data_bar,
        conditional_formatting_additional::add_top_bottom_rule,
        conditional_formatting_additional::add_text_rule,
        conditional_formatting_additional::add_icon_set,
        conditional_formatting_additional::add_above_below_average_rule,
        // CSV functions
        write_csv_with_options::write_csv,
        write_csv_with_options::write_csv_with_options,
        // Background color functions
        set_background_color::set_background_color,
        // Cell alignment functions
        set_cell_alignment::set_cell_alignment,
        // Data validation functions
        data_validation::add_number_validation,
        data_validation::add_date_validation,
        data_validation::add_text_length_validation,
        data_validation::add_list_validation,
        data_validation::add_custom_validation,
        data_validation::remove_data_validation,
        // Sheet view functions
        sheet_view_functions::set_show_grid_lines,
        sheet_view_functions::set_tab_selected,
        sheet_view_functions::set_top_left_cell,
        sheet_view_functions::set_zoom_scale,
        sheet_view_functions::freeze_panes,
        sheet_view_functions::split_panes,
        sheet_view_functions::set_tab_color,
        sheet_view_functions::set_sheet_view,
        sheet_view_functions::set_zoom_scale_normal,
        sheet_view_functions::set_zoom_scale_page_layout,
        sheet_view_functions::set_zoom_scale_page_break,
        sheet_view_functions::set_selection,
        sheet_view_functions::get_show_grid_lines,
        sheet_view_functions::get_zoom_scale,
        sheet_view_functions::get_tab_color,
        sheet_view_functions::get_sheet_view,
        sheet_view_functions::get_selection,
        // Workbook functions
        workbook_view_functions::get_active_tab,
        workbook_view_functions::get_workbook_window_position,
        workbook_view_functions::set_active_tab,
        workbook_view_functions::set_workbook_window_position,
        workbook_protection_functions::is_workbook_protected,
        workbook_protection_functions::get_workbook_protection_details,
        workbook_protection_functions::set_workbook_protection,
        // Comment functions
        comment_functions::add_comment,
        comment_functions::get_comment,
        comment_functions::update_comment,
        comment_functions::remove_comment,
        comment_functions::has_comments,
        comment_functions::get_comments_count,
        // File format functions
        file_format_options::write_with_compression,
        file_format_options::write_with_encryption_options,
        file_format_options::to_binary_xlsx,
        // Password functions
        password_functions::set_password,
        // Formula functions
        formula_functions::set_formula,
        formula_functions::set_array_formula,
        formula_functions::create_named_range,
        formula_functions::create_defined_name,
        formula_functions::get_defined_names,
        // Formula getter functions
        formula_functions::is_formula,
        formula_functions::get_formula,
        formula_functions::get_formula_obj,
        formula_functions::get_formula_shared_index,
        formula_functions::get_text,
        formula_functions::get_formula_type,
        formula_functions::get_shared_index,
        formula_functions::get_reference,
        formula_functions::get_bx,
        formula_functions::get_data_table_2d,
        formula_functions::get_data_table_row,
        formula_functions::get_input_1deleted,
        formula_functions::get_input_2deleted,
        formula_functions::get_r1,
        formula_functions::get_r2,
        // Auto filter functions
        auto_filter_functions::set_auto_filter,
        auto_filter_functions::remove_auto_filter,
        auto_filter_functions::has_auto_filter,
        auto_filter_functions::get_auto_filter_range,
        // Table functions
        table::add_table,
        table::get_tables,
        table::get_table,
        table::remove_table,
        table::has_tables,
        table::count_tables,
        table::set_table_style,
        table::get_table_style,
        table::remove_table_style,
        table::add_table_column,
        table::get_table_columns,
        table::modify_table_column,
        table::set_table_totals_row,
        table::get_table_totals_row,
        // Hyperlink functions
        hyperlink::add_hyperlink,
        hyperlink::get_hyperlink,
        hyperlink::remove_hyperlink,
        hyperlink::has_hyperlink,
        hyperlink::has_hyperlinks,
        hyperlink::get_hyperlinks,
        hyperlink::update_hyperlink,
        // Rich text functions
        rich_text_functions::create_rich_text,
        rich_text_functions::create_rich_text_from_html,
        rich_text_functions::create_text_element,
        rich_text_functions::get_text_element_text,
        rich_text_functions::get_text_element_font_properties,
        rich_text_functions::add_text_element_to_rich_text,
        rich_text_functions::add_formatted_text_to_rich_text,
        rich_text_functions::set_cell_rich_text,
        rich_text_functions::get_cell_rich_text,
        rich_text_functions::get_rich_text_plain_text,
        rich_text_functions::rich_text_to_html,
        rich_text_functions::get_rich_text_elements,
        // OLE Objects functions
        ole_object_functions::create_ole_objects,
        ole_object_functions::create_ole_object,
        ole_object_functions::create_embedded_object_properties,
        ole_object_functions::get_ole_objects,
        ole_object_functions::set_ole_objects,
        ole_object_functions::add_ole_object,
        ole_object_functions::get_ole_object_list,
        ole_object_functions::count_ole_objects,
        ole_object_functions::has_ole_objects,
        ole_object_functions::get_ole_object_requires,
        ole_object_functions::set_ole_object_requires,
        ole_object_functions::get_ole_object_prog_id,
        ole_object_functions::set_ole_object_prog_id,
        ole_object_functions::get_ole_object_extension,
        ole_object_functions::set_ole_object_extension,
        ole_object_functions::get_ole_object_data,
        ole_object_functions::set_ole_object_data,
        ole_object_functions::get_ole_object_properties,
        ole_object_functions::set_ole_object_properties,
        ole_object_functions::get_embedded_object_prog_id,
        ole_object_functions::set_embedded_object_prog_id,
        ole_object_functions::get_embedded_object_shape_id,
        ole_object_functions::set_embedded_object_shape_id,
        ole_object_functions::load_ole_object_from_file,
        ole_object_functions::save_ole_object_to_file,
        ole_object_functions::create_ole_object_with_data,
        ole_object_functions::is_ole_object_binary,
        ole_object_functions::is_ole_object_excel,
        // Page breaks functions
        page_breaks::add_row_page_break,
        page_breaks::add_column_page_break,
        page_breaks::remove_row_page_break,
        page_breaks::remove_column_page_break,
        page_breaks::get_row_page_breaks,
        page_breaks::get_column_page_breaks,
        page_breaks::clear_row_page_breaks,
        page_breaks::clear_column_page_breaks,
        page_breaks::has_row_page_break,
        page_breaks::has_column_page_break,
        // Document Properties functions
        document_properties::get_custom_property,
        document_properties::set_custom_property_string,
        document_properties::set_custom_property_number,
        document_properties::set_custom_property_bool,
        document_properties::set_custom_property_date,
        document_properties::remove_custom_property,
        document_properties::get_custom_property_names,
        document_properties::has_custom_property,
        document_properties::get_custom_properties_count,
        document_properties::clear_custom_properties,
        document_properties::get_title,
        document_properties::set_title,
        document_properties::get_description,
        document_properties::set_description,
        document_properties::get_subject,
        document_properties::set_subject,
        document_properties::get_keywords,
        document_properties::set_keywords,
        document_properties::get_creator,
        document_properties::set_creator,
        document_properties::get_last_modified_by,
        document_properties::set_last_modified_by,
        document_properties::get_category,
        document_properties::set_category,
        document_properties::get_company,
        document_properties::set_company,
        document_properties::get_manager,
        document_properties::set_manager,
        document_properties::get_created,
        document_properties::set_created,
        document_properties::get_modified,
        document_properties::set_modified,
        // VML support functions
        vml_support::create_vml_shape,
        vml_support::set_vml_shape_style,
        vml_support::set_vml_shape_type,
        vml_support::set_vml_shape_filled,
        vml_support::set_vml_shape_fill_color,
        vml_support::set_vml_shape_stroked,
        vml_support::set_vml_shape_stroke_color,
        vml_support::set_vml_shape_stroke_weight,
        // File Format Options functions
        file_format_options::write_with_compression,
        file_format_options::write_with_encryption_options,
        file_format_options::to_binary_xlsx,
        file_format_options::get_compression_level,
        file_format_options::is_encrypted,
        file_format_options::get_encryption_algorithm,
    ]
);
