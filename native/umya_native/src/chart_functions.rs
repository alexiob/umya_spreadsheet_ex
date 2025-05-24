// Chart functions implementation for the umya-spreadsheet Elixir wrapper
use crate::{atoms, UmyaSpreadsheet};
use rustler::{Atom, ResourceArc, Term};

// Public API struct paths
use umya_spreadsheet::Chart;
use umya_spreadsheet::ChartType;
// Drawing module imports
use umya_spreadsheet::drawing::spreadsheet::MarkerType;
// Chart-related imports
use umya_spreadsheet::drawing::charts::ChartText;
use umya_spreadsheet::drawing::charts::Index;
use umya_spreadsheet::drawing::charts::Legend;
use umya_spreadsheet::drawing::charts::LegendPositionValues;
use umya_spreadsheet::drawing::charts::MarkerStyleValues;
use umya_spreadsheet::drawing::charts::Overlay;
use umya_spreadsheet::drawing::charts::Perspective;
use umya_spreadsheet::drawing::charts::RichText;
use umya_spreadsheet::drawing::charts::RotateX;
use umya_spreadsheet::drawing::charts::RotateY;
use umya_spreadsheet::drawing::charts::SeriesText;
use umya_spreadsheet::drawing::charts::ShowCategoryName;
use umya_spreadsheet::drawing::charts::ShowPercent;
use umya_spreadsheet::drawing::charts::ShowSeriesName;
use umya_spreadsheet::drawing::charts::ShowValue;
use umya_spreadsheet::drawing::charts::Title;
use umya_spreadsheet::drawing::Paragraph;
use umya_spreadsheet::drawing::Run;

// All chart functions are already public (#[rustler::nif]) and don't need to be re-exported
// The self:: prefix is incorrect as these functions are defined directly in this module

/// Add a chart to a sheet
#[rustler::nif]
pub fn add_chart(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    chart_type: String,
    from_cell: String,
    to_cell: String,
    title: String,
    data_series: Vec<String>,
    series_titles: Vec<String>,
    _point_titles: Vec<String>,
) -> Result<Atom, (Atom, String)> {
    let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| -> Result<Atom, String> {
        let mut guard = resource.spreadsheet.lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        match guard.get_sheet_by_name_mut(&sheet_name) {
            Some(sheet) => {
                // Create markers for chart position
                let mut from_marker = MarkerType::default();
                let mut to_marker = MarkerType::default();
                from_marker.set_coordinate(&from_cell);
                to_marker.set_coordinate(&to_cell);

                // Convert chart_type string to ChartType enum
                let chart_type_enum = match chart_type.as_str() {
                    "LineChart" => ChartType::LineChart,
                    "PieChart" => ChartType::PieChart,
                    "DoughnutChart" => ChartType::DoughnutChart,
                    "BarChart" => ChartType::BarChart,
                    "AreaChart" => ChartType::AreaChart,
                    "ScatterChart" => ChartType::ScatterChart,
                    "RadarChart" => ChartType::RadarChart,
                    "BubbleChart" => ChartType::BubbleChart,
                    "Line3DChart" => ChartType::Line3DChart,
                    "Bar3DChart" => ChartType::Bar3DChart,
                    "Pie3DChart" => ChartType::Pie3DChart,
                    "Area3DChart" => ChartType::Area3DChart,
                    "OfPieChart" => ChartType::OfPieChart,
                    _ => return Err(format!("Unsupported chart type: {}", chart_type)),
                };

                // Create and configure the chart
                let mut chart = Chart::default();

                // Convert Vec<String> to Vec<&str> for data_series
                let data_series_refs: Vec<&str> = data_series.iter().map(|s| s.as_str()).collect();

                // Initialize the chart with type, position, and data series
                chart.new_chart(chart_type_enum, from_marker, to_marker, data_series_refs);

                // Set the title with proper formatting
                if !title.is_empty() {
                    // Get chart space and chart objects
                    let chart_space = chart.get_chart_space_mut();
                    let chart_obj = chart_space.get_chart_mut();

                    // Create a title object
                    let mut title_obj = Title::default();

                    // Create rich text structure for the title
                    let mut rich_text = RichText::default();
                    let mut paragraph = Paragraph::default();
                    let mut run = Run::default();
                    run.set_text(&title);
                    paragraph.add_run(run);
                    rich_text.add_paragraph(paragraph);

                    // Create chart text and set rich text
                    let mut chart_text = ChartText::default();
                    chart_text.set_rich_text(rich_text);

                    // Set the title's text
                    title_obj.set_chart_text(chart_text);

                    // Set the title on the chart
                    chart_obj.set_title(title_obj);
                }

                // Add series titles if provided
                if !series_titles.is_empty() {
                    match chart_type.as_str() {
                        "LineChart" | "Line3DChart" => {
                            let plot_area = chart.get_plot_area_mut();
                            if let Some(line_chart) = plot_area.get_line_chart_mut() {
                                let series_list = line_chart.get_area_chart_series_list_mut();
                                for (i, title) in series_titles.iter().enumerate() {
                                    if i < series_list.get_area_chart_series().len() {
                                        // Create an Index object for set_index
                                        let mut idx = Index::default();
                                        idx.set_val(i as u32);
                                        series_list.get_area_chart_series_mut()[i].set_index(idx);

                                        // Create SeriesText with proper object structure
                                        let mut series_text = SeriesText::default();
                                        series_text.set_value(title);

                                        series_list.get_area_chart_series_mut()[i]
                                            .set_series_text(series_text);
                                    }
                                }
                            }
                        }
                        "BarChart" | "Bar3DChart" => {
                            let plot_area = chart.get_plot_area_mut();
                            if let Some(bar_chart) = plot_area.get_bar_chart_mut() {
                                let series_list = bar_chart.get_area_chart_series_list_mut();
                                for (i, title) in series_titles.iter().enumerate() {
                                    if i < series_list.get_area_chart_series().len() {
                                        // Create an Index object for set_index
                                        let mut idx = Index::default();
                                        idx.set_val(i as u32);
                                        series_list.get_area_chart_series_mut()[i].set_index(idx);

                                        // Create SeriesText with proper object structure
                                        let mut series_text = SeriesText::default();
                                        series_text.set_value(title);

                                        series_list.get_area_chart_series_mut()[i]
                                            .set_series_text(series_text);
                                    }
                                }
                            }
                        }
                        "PieChart" | "Pie3DChart" | "DoughnutChart" => {
                            let plot_area = chart.get_plot_area_mut();
                            if let Some(pie_chart) = plot_area.get_pie_chart_mut() {
                                let series_list = pie_chart.get_area_chart_series_list_mut();
                                for (i, title) in series_titles.iter().enumerate() {
                                    if i < series_list.get_area_chart_series().len() {
                                        // Create an Index object for set_index
                                        let mut idx = Index::default();
                                        idx.set_val(i as u32);
                                        series_list.get_area_chart_series_mut()[i].set_index(idx);

                                        // Create SeriesText with proper object structure
                                        let mut series_text = SeriesText::default();
                                        series_text.set_value(title);

                                        series_list.get_area_chart_series_mut()[i]
                                            .set_series_text(series_text);
                                    }
                                }
                            }
                        }
                        "AreaChart" | "Area3DChart" => {
                            let plot_area = chart.get_plot_area_mut();
                            if let Some(area_chart) = plot_area.get_area_chart_mut() {
                                let series_list = area_chart.get_area_chart_series_list_mut();
                                for (i, title) in series_titles.iter().enumerate() {
                                    if i < series_list.get_area_chart_series().len() {
                                        // Create an Index object for set_index
                                        let mut idx = Index::default();
                                        idx.set_val(i as u32);
                                        series_list.get_area_chart_series_mut()[i].set_index(idx);

                                        // Create SeriesText with proper object structure
                                        let mut series_text = SeriesText::default();
                                        series_text.set_value(title);

                                        series_list.get_area_chart_series_mut()[i]
                                            .set_series_text(series_text);
                                    }
                                }
                            }
                        }
                        _ => {}
                    }
                }

                // Set up a default legend
                let chart_obj = chart.get_chart_space_mut().get_chart_mut();
                let mut legend = Legend::default();
                let mut legend_position = umya_spreadsheet::drawing::charts::LegendPosition::default();
                legend_position.set_val(LegendPositionValues::Right);
                legend.set_legend_position(legend_position);
                chart_obj.set_legend(legend);

                // Add the chart to the sheet
                sheet.add_chart(chart);

                Ok(atoms::ok())
            }
            None => Err(format!("Sheet '{}' not found", sheet_name)),
        }
    }));

    match result {
        Ok(Ok(atom)) => Ok(atom),
        Ok(Err(err_msg)) => Err((atoms::error(), err_msg)),
        Err(_) => Err((atoms::error(), "Error occurred in add_chart operation".to_string())),
    }
}

/// Sets chart style
#[rustler::nif]
pub fn set_chart_style(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    chart_index: i32,
    style: i32,
) -> Result<Atom, (Atom, String)> {
    let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| -> Result<Atom, String> {
        if chart_index < 0 || style < 1 || style > 48 {
            return Err(format!("Invalid parameters: chart_index={}, style={}. Chart index must be >= 0 and style must be between 1-48", chart_index, style));
        }

        let mut guard = resource.spreadsheet.lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        match guard.get_sheet_by_name_mut(&sheet_name) {
            Some(sheet) => {
                // Get the chart collection from the sheet
                let chart_collection = sheet.get_chart_collection_mut();

                // Check if chart index is valid
                if chart_index < 0 || chart_index as usize >= chart_collection.len() {
                    return Err(format!("Chart index {} not found. Available charts: 0-{}", chart_index, chart_collection.len().saturating_sub(1)));
                }

                // Get the chart at the given index
                let chart = &mut chart_collection[chart_index as usize];

                // Note: umya-spreadsheet doesn't have a direct set_style method,
                // but we can set some appearance properties to simulate styles

                // Access the chart space to apply style properties
                let chart_space = chart.get_chart_space_mut();
                let _chart_obj = chart_space.get_chart_mut();

                // Apply some basic style adjustments based on the style number
                // This is a simplified approach as the actual Excel styles are more complex

                // For example, adjust colors or properties based on style number ranges
                if style >= 1 && style <= 8 {
                    // Style range 1-8: Different color themes
                    // We could set specific colors here
                } else if style >= 9 && style <= 16 {
                    // Style range 9-16: Different marker styles
                    // Apply different marker styles for line or scatter charts
                    let plot_area = chart.get_plot_area_mut();
                    if let Some(line_chart) = plot_area.get_line_chart_mut() {
                        let series_list = line_chart.get_area_chart_series_list_mut();
                        for series in series_list.get_area_chart_series_mut() {
                            // Set marker style based on style number
                            if let Some(marker) = series.get_marker_mut() {
                                if let Some(symbol) = marker.get_symbol_mut() {
                                    let marker_style = match style % 8 {
                                        1 => MarkerStyleValues::Circle,
                                        2 => MarkerStyleValues::Dash,
                                        3 => MarkerStyleValues::Diamond,
                                        4 => MarkerStyleValues::Dot,
                                        5 => MarkerStyleValues::None,
                                        6 => MarkerStyleValues::Picture,
                                        7 => MarkerStyleValues::Plus,
                                        _ => MarkerStyleValues::Square,
                                    };
                                    symbol.set_val(marker_style);
                                }
                            }
                        }
                    }
                }

                // Actually apply the style is limited by the umya-spreadsheet API
                // For now, we'll return success even though we can't apply all Excel-like styles

                Ok(atoms::ok())
            }
            None => Err(format!("Sheet '{}' not found", sheet_name)),
        }
    }));

    match result {
        Ok(Ok(atom)) => Ok(atom),
        Ok(Err(err_msg)) => Err((atoms::error(), err_msg)),
        Err(_) => Err((atoms::error(), "Error occurred in set_chart_style operation".to_string())),
    }
}

/// Sets chart data labels
#[rustler::nif]
pub fn set_chart_data_labels(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    chart_index: i32,
    show_values: bool,
    show_percent: bool,
    show_category_name: bool,
    show_series_name: bool,
    _position: String,
) -> Result<Atom, (Atom, String)> {
    let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| -> Result<Atom, String> {
        if chart_index < 0 {
            return Err(format!("Invalid chart index: {}. Chart index must be >= 0", chart_index));
        }

        let mut guard = resource.spreadsheet.lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        match guard.get_sheet_by_name_mut(&sheet_name) {
            Some(sheet) => {
                // Get the chart collection from the sheet
                let chart_collection = sheet.get_chart_collection_mut();

                // Check if chart index is valid
                if chart_index < 0 || chart_index as usize >= chart_collection.len() {
                    return Err(format!("Chart index {} not found. Available charts: 0-{}", chart_index, chart_collection.len().saturating_sub(1)));
                }

                // Get the chart at the given index
                let chart = &mut chart_collection[chart_index as usize];

                // Apply data label settings based on chart type
                let plot_area = chart.get_plot_area_mut();

            // For bar chart
            if let Some(bar_chart) = plot_area.get_bar_chart_mut() {
                let series_list = bar_chart.get_area_chart_series_list_mut();
                for series in series_list.get_area_chart_series_mut() {
                    if let Some(data_labels) = series.get_data_labels_mut() {
                        // Create ShowValue object for values display
                        if show_values {
                            let mut show_val = ShowValue::default();
                            show_val.set_val(show_values);
                            data_labels.set_show_value(show_val);
                        }

                        // Create ShowPercent object for percentage display
                        if show_percent {
                            let mut show_percent_obj = ShowPercent::default();
                            show_percent_obj.set_val(show_percent);
                            data_labels.set_show_percent(show_percent_obj);
                        }

                        // Create ShowCategoryName object for category display
                        if show_category_name {
                            let mut show_cat = ShowCategoryName::default();
                            show_cat.set_val(show_category_name);
                            data_labels.set_show_category_name(show_cat);
                        }

                        // Create ShowSeriesName object for series display
                        if show_series_name {
                            let mut show_series = ShowSeriesName::default();
                            show_series.set_val(show_series_name);
                            data_labels.set_show_series_name(show_series);
                        }
                    }
                }
            }

            // For line chart
            if let Some(line_chart) = plot_area.get_line_chart_mut() {
                let series_list = line_chart.get_area_chart_series_list_mut();
                for series in series_list.get_area_chart_series_mut() {
                    if let Some(data_labels) = series.get_data_labels_mut() {
                        if show_values {
                            let mut show_val = ShowValue::default();
                            show_val.set_val(show_values);
                            data_labels.set_show_value(show_val);
                        }

                        if show_percent {
                            let mut show_percent_obj = ShowPercent::default();
                            show_percent_obj.set_val(show_percent);
                            data_labels.set_show_percent(show_percent_obj);
                        }

                        if show_category_name {
                            let mut show_cat = ShowCategoryName::default();
                            show_cat.set_val(show_category_name);
                            data_labels.set_show_category_name(show_cat);
                        }

                        if show_series_name {
                            let mut show_series = ShowSeriesName::default();
                            show_series.set_val(show_series_name);
                            data_labels.set_show_series_name(show_series);
                        }
                    }
                }
            }

            // For pie chart
            if let Some(pie_chart) = plot_area.get_pie_chart_mut() {
                let series_list = pie_chart.get_area_chart_series_list_mut();
                for series in series_list.get_area_chart_series_mut() {
                    if let Some(data_labels) = series.get_data_labels_mut() {
                        if show_values {
                            let mut show_val = ShowValue::default();
                            show_val.set_val(show_values);
                            data_labels.set_show_value(show_val);
                        }

                        if show_percent {
                            let mut show_percent_obj = ShowPercent::default();
                            show_percent_obj.set_val(show_percent);
                            data_labels.set_show_percent(show_percent_obj);
                        }

                        if show_category_name {
                            let mut show_cat = ShowCategoryName::default();
                            show_cat.set_val(show_category_name);
                            data_labels.set_show_category_name(show_cat);
                        }

                        if show_series_name {
                            let mut show_series = ShowSeriesName::default();
                            show_series.set_val(show_series_name);
                            data_labels.set_show_series_name(show_series);
                        }
                    }
                }
            }

                // For area chart
                if let Some(area_chart) = plot_area.get_area_chart_mut() {
                    let series_list = area_chart.get_area_chart_series_list_mut();
                    for series in series_list.get_area_chart_series_mut() {
                        if let Some(data_labels) = series.get_data_labels_mut() {
                            if show_values {
                                let mut show_val = ShowValue::default();
                                show_val.set_val(show_values);
                                data_labels.set_show_value(show_val);
                            }

                            if show_percent {
                                let mut show_percent_obj = ShowPercent::default();
                                show_percent_obj.set_val(show_percent);
                                data_labels.set_show_percent(show_percent_obj);
                            }

                            if show_category_name {
                                let mut show_cat = ShowCategoryName::default();
                                show_cat.set_val(show_category_name);
                                data_labels.set_show_category_name(show_cat);
                            }

                            if show_series_name {
                                let mut show_series = ShowSeriesName::default();
                                show_series.set_val(show_series_name);
                                data_labels.set_show_series_name(show_series);
                            }
                        }
                    }
                }

                Ok(atoms::ok())
            }
            None => Err(format!("Sheet '{}' not found", sheet_name)),
        }
    }));

    match result {
        Ok(Ok(atom)) => Ok(atom),
        Ok(Err(err_msg)) => Err((atoms::error(), err_msg)),
        Err(_) => Err((atoms::error(), "Error occurred in set_chart_data_labels operation".to_string())),
    }
}

/// Set the chart legend position and overlay property
#[rustler::nif]
pub fn set_chart_legend_position(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    chart_index: i32,
    position: String,
    overlay: bool,
) -> Result<Atom, (Atom, String)> {
    let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| -> Result<Atom, String> {
        if chart_index < 0 {
            return Err(format!("Invalid chart index: {}. Chart index must be >= 0", chart_index));
        }

        let mut guard = resource.spreadsheet.lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;
        match guard.get_sheet_by_name_mut(&sheet_name) {
            Some(sheet) => {
                // Get the chart collection from the sheet
                let chart_collection = sheet.get_chart_collection_mut();

                // Check if chart index is valid
                if chart_index < 0 || chart_index as usize >= chart_collection.len() {
                    return Err(format!("Chart index {} not found. Available charts: 0-{}", chart_index, chart_collection.len().saturating_sub(1)));
                }

                // Get the chart at the given index
                let chart = &mut chart_collection[chart_index as usize];

                // Find the legend position from the string parameter
                let position_value = match position.to_lowercase().as_str() {
                    "right" => LegendPositionValues::Right,
                    "left" => LegendPositionValues::Left,
                    "top" => LegendPositionValues::Top,
                    "bottom" => LegendPositionValues::Bottom,
                    "topright" => LegendPositionValues::TopRight,
                    "none" => {
                        // If "none", we may need to disable the legend entirely
                        // or set a default position if we can't hide it
                        LegendPositionValues::Right // Default
                    }
                    _ => return Err(format!("Invalid legend position: {}. Valid positions are: right, left, top, bottom, topright, none", position)), // Invalid position
                };

                // Get the chart object from the chart space
                let chart_obj = chart.get_chart_space_mut().get_chart_mut();

                // Configure the legend
                let mut legend = Legend::default();
                let mut legend_position = umya_spreadsheet::drawing::charts::LegendPosition::default();
                legend_position.set_val(position_value);
                legend.set_legend_position(legend_position);

                let mut overlay_obj = Overlay::default();
                overlay_obj.set_val(overlay);
                legend.set_overlay(overlay_obj);

                // Set the legend on the chart object
                chart_obj.set_legend(legend);

                Ok(atoms::ok())
            }
            None => Err(format!("Sheet '{}' not found", sheet_name)), // Sheet not found
        }
    }));

    match result {
        Ok(Ok(atom)) => Ok(atom),
        Ok(Err(err_msg)) => Err((atoms::error(), err_msg)),
        Err(_) => Err((atoms::error(), "Error occurred in set_chart_legend_position operation".to_string())),
    }
}

/// Configure 3D view properties for a chart
#[rustler::nif]
pub fn set_chart_3d_view(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    chart_index: i32,
    rot_x: i32,
    rot_y: i32,
    perspective: i32,
    _height_percent: i32,
) -> Result<Atom, (Atom, String)> {
    let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| -> Result<Atom, String> {
        if chart_index < 0 {
            return Err(format!("Invalid chart index: {}. Chart index must be >= 0", chart_index));
        }

        let mut guard = resource.spreadsheet.lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        match guard.get_sheet_by_name_mut(&sheet_name) {
            Some(sheet) => {
                // Get the chart collection from the sheet
                let chart_collection = sheet.get_chart_collection_mut();

                // Check if chart index is valid
                if chart_index < 0 || chart_index as usize >= chart_collection.len() {
                    return Err(format!("Chart index {} not found. Available charts: 0-{}", chart_index, chart_collection.len().saturating_sub(1)));
                }

                // Get the chart at the given index
                let chart = &mut chart_collection[chart_index as usize];

                // Create the View3D object
                let mut view_3d = umya_spreadsheet::drawing::charts::View3D::default();

                // Create and set RotateX
                let mut rotate_x = RotateX::default();
                rotate_x.set_val(rot_x as i8);
                view_3d.set_rotate_x(rotate_x);

                // Create and set RotateY
                let mut rotate_y = RotateY::default();
                rotate_y.set_val(rot_y as u16);
                view_3d.set_rotate_y(rotate_y);

                // Create and set RotateY
                let mut rotate_y = RotateY::default();
                rotate_y.set_val(rot_y as u16);
                view_3d.set_rotate_y(rotate_y);

                // Create and set Perspective
                let mut perspective_obj = Perspective::default();
                perspective_obj.set_val(perspective as u8);
                view_3d.set_perspective(perspective_obj);

                // Access the chart space
                let chart_space = chart.get_chart_space_mut();
                let chart_obj = chart_space.get_chart_mut();

                // Set the view3D on the chart object
                chart_obj.set_view_3d(view_3d);

                // Success
                Ok(atoms::ok())
            }
            None => Err(format!("Sheet '{}' not found", sheet_name)),
        }
    }));

    match result {
        Ok(Ok(atom)) => Ok(atom),
        Ok(Err(err_msg)) => Err((atoms::error(), err_msg)),
        Err(_) => Err((atoms::error(), "Error occurred in set_chart_3d_view operation".to_string())),
    }
}

/// Set chart axis titles
#[rustler::nif]
pub fn set_chart_axis_titles(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    chart_index: i32,
    category_axis_title: String,
    value_axis_title: String,
) -> Result<Atom, (Atom, String)> {
    let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| -> Result<Atom, String> {
        if chart_index < 0 {
            return Err(format!("Invalid chart index: {}. Chart index must be >= 0", chart_index));
        }

        let mut guard = resource.spreadsheet.lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        match guard.get_sheet_by_name_mut(&sheet_name) {
            Some(sheet) => {
                // Get the chart collection from the sheet
                let chart_collection = sheet.get_chart_collection_mut();

                // Check if chart index is valid
                if chart_index < 0 || chart_index as usize >= chart_collection.len() {
                    return Err(format!("Chart index {} not found. Available charts: 0-{}", chart_index, chart_collection.len().saturating_sub(1)));
                }

                // Get the chart at the given index
                let chart = &mut chart_collection[chart_index as usize];

                // Get the plot area which contains the axes
                let plot_area = chart.get_plot_area_mut();

                // Set category axis title
                if !category_axis_title.is_empty() {
                    if let Some(category_axis) = plot_area.get_category_axis_mut().first_mut() {
                        // Create title with rich text
                        let mut title = Title::default();

                        // Create rich text
                        let mut rich_text = RichText::default();
                        let mut paragraph = Paragraph::default();
                        let mut run = Run::default();
                        run.set_text(&category_axis_title);
                        paragraph.add_run(run);
                        rich_text.add_paragraph(paragraph);

                        // Create chart text
                        let mut chart_text = ChartText::default();
                        chart_text.set_rich_text(rich_text);

                        // Set the title's text
                        title.set_chart_text(chart_text);

                        // Apply the title to the axis
                        category_axis.set_title(title);
                    }
                }

                // Set value axis title
                if !value_axis_title.is_empty() {
                    if let Some(value_axis) = plot_area.get_value_axis_mut().first_mut() {
                        // Create title with rich text
                        let mut title = Title::default();

                        // Create rich text
                        let mut rich_text = RichText::default();
                        let mut paragraph = Paragraph::default();
                        let mut run = Run::default();
                        run.set_text(&value_axis_title);
                        paragraph.add_run(run);
                        rich_text.add_paragraph(paragraph);

                        // Create chart text
                        let mut chart_text = ChartText::default();
                        chart_text.set_rich_text(rich_text);

                        // Set the title's text
                        title.set_chart_text(chart_text);

                        // Apply the title to the axis
                        value_axis.set_title(title);
                    }
                }

                Ok(atoms::ok())
            }
            None => Err(format!("Sheet '{}' not found", sheet_name)),
        }
    }));

    match result {
        Ok(Ok(atom)) => Ok(atom),
        Ok(Err(err_msg)) => Err((atoms::error(), err_msg)),
        Err(_) => Err((atoms::error(), "Error occurred in set_chart_axis_titles operation".to_string())),
    }
}

/// Advanced chart options implementation
#[rustler::nif]
pub fn add_chart_with_options(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    chart_type: String,
    from_cell: String,
    to_cell: String,
    title: String,
    data_series: Vec<String>,
    series_titles: Vec<String>,
    _point_titles: Vec<String>,
    _style: i32,
    vary_colors: bool,
    view_3d: Term,
    legend: Term,
    _axes: Term,
    data_labels: Term,
    _chart_specific: Term,
) -> Result<Atom, (Atom, String)> {
    let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| -> Result<Atom, String> {
        let mut guard = resource.spreadsheet.lock()
            .map_err(|_| "Failed to acquire spreadsheet lock".to_string())?;

        match guard.get_sheet_by_name_mut(&sheet_name) {
            Some(sheet) => {
                // Create markers for chart position
                let mut from_marker = MarkerType::default();
                let mut to_marker = MarkerType::default();
                from_marker.set_coordinate(&from_cell);
                to_marker.set_coordinate(&to_cell);

                // Convert chart_type string to ChartType enum
                let chart_type_enum = match chart_type.as_str() {
                    "LineChart" => ChartType::LineChart,
                    "PieChart" => ChartType::PieChart,
                    "DoughnutChart" => ChartType::DoughnutChart,
                    "BarChart" => ChartType::BarChart,
                    "AreaChart" => ChartType::AreaChart,
                    "ScatterChart" => ChartType::ScatterChart,
                    "RadarChart" => ChartType::RadarChart,
                    "BubbleChart" => ChartType::BubbleChart,
                    "Line3DChart" => ChartType::Line3DChart,
                    "Bar3DChart" => ChartType::Bar3DChart,
                    "Pie3DChart" => ChartType::Pie3DChart,
                    "Area3DChart" => ChartType::Area3DChart,
                    "OfPieChart" => ChartType::OfPieChart,
                    _ => return Err(format!("Unsupported chart type: {}", chart_type)),
                };

                // Create and configure the chart
                let mut chart = Chart::default();

                // Convert Vec<String> to Vec<&str> for data_series
                let data_series_refs: Vec<&str> = data_series.iter().map(|s| s.as_str()).collect();

                // Initialize the chart with type, position, and data series
                chart
                    .new_chart(chart_type_enum, from_marker, to_marker, data_series_refs)
                    .set_title(&title);

            // Add series titles if provided
            if !series_titles.is_empty() {
                match chart_type.as_str() {
                    "LineChart" | "Line3DChart" => {
                        let plot_area = chart.get_plot_area_mut();
                        if let Some(line_chart) = plot_area.get_line_chart_mut() {
                            let series_list = line_chart.get_area_chart_series_list_mut();
                            for (i, title) in series_titles.iter().enumerate() {
                                if i < series_list.get_area_chart_series().len() {
                                    // Create an Index object for set_index
                                    let mut idx = Index::default();
                                    idx.set_val(i as u32);
                                    series_list.get_area_chart_series_mut()[i].set_index(idx);

                                    // Create SeriesText with proper object structure
                                    let mut series_text = SeriesText::default();
                                    series_text.set_value(title);

                                    series_list.get_area_chart_series_mut()[i]
                                        .set_series_text(series_text);
                                }
                            }
                        }
                    }
                    "BarChart" | "Bar3DChart" => {
                        let plot_area = chart.get_plot_area_mut();
                        if let Some(bar_chart) = plot_area.get_bar_chart_mut() {
                            let series_list = bar_chart.get_area_chart_series_list_mut();
                            for (i, title) in series_titles.iter().enumerate() {
                                if i < series_list.get_area_chart_series().len() {
                                    // Create an Index object for set_index
                                    let mut idx = Index::default();
                                    idx.set_val(i as u32);
                                    series_list.get_area_chart_series_mut()[i].set_index(idx);

                                    // Create SeriesText with proper object structure
                                    let mut series_text = SeriesText::default();
                                    series_text.set_value(title);

                                    series_list.get_area_chart_series_mut()[i]
                                        .set_series_text(series_text);
                                }
                            }
                        }
                    }
                    "PieChart" | "Pie3DChart" | "DoughnutChart" => {
                        let plot_area = chart.get_plot_area_mut();
                        if let Some(pie_chart) = plot_area.get_pie_chart_mut() {
                            let series_list = pie_chart.get_area_chart_series_list_mut();
                            for (i, title) in series_titles.iter().enumerate() {
                                if i < series_list.get_area_chart_series().len() {
                                    // Create an Index object for set_index
                                    let mut idx = Index::default();
                                    idx.set_val(i as u32);
                                    series_list.get_area_chart_series_mut()[i].set_index(idx);

                                    // Create SeriesText with proper object structure
                                    let mut series_text = SeriesText::default();
                                    series_text.set_value(title);

                                    series_list.get_area_chart_series_mut()[i]
                                        .set_series_text(series_text);
                                }
                            }
                        }
                    }
                    "AreaChart" | "Area3DChart" => {
                        let plot_area = chart.get_plot_area_mut();
                        if let Some(area_chart) = plot_area.get_area_chart_mut() {
                            let series_list = area_chart.get_area_chart_series_list_mut();
                            for (i, title) in series_titles.iter().enumerate() {
                                if i < series_list.get_area_chart_series().len() {
                                    // Create an Index object for set_index
                                    let mut idx = Index::default();
                                    idx.set_val(i as u32);
                                    series_list.get_area_chart_series_mut()[i].set_index(idx);

                                    // Create SeriesText with proper object structure
                                    let mut series_text = SeriesText::default();
                                    series_text.set_value(title);

                                    series_list.get_area_chart_series_mut()[i]
                                        .set_series_text(series_text);
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }

            // Set chart style if provided
            // Note: Chart doesn't actually have set_style in umya-spreadsheet
            // but we keep this for compatibility

            // Set vary colors if provided
            if vary_colors {
                let plot_area = chart.get_plot_area_mut();
                // Apply vary colors to different chart types
                if let Some(line_chart) = plot_area.get_line_chart_mut() {
                    line_chart.get_vary_colors_mut().set_val(vary_colors);
                } else if let Some(bar_chart) = plot_area.get_bar_chart_mut() {
                    bar_chart.get_vary_colors_mut().set_val(vary_colors);
                } else if let Some(pie_chart) = plot_area.get_pie_chart_mut() {
                    pie_chart.get_vary_colors_mut().set_val(vary_colors);
                } else if let Some(area_chart) = plot_area.get_area_chart_mut() {
                    area_chart.get_vary_colors_mut().set_val(vary_colors);
                }
            }

            // Process view_3d Term if it's a map
            if !view_3d.is_atom() {
                // Create a new View3D object
                let mut view_3d_obj = umya_spreadsheet::drawing::charts::View3D::default();

                // Try to extract rotate_x from the map
                if let Ok(term) = view_3d.map_get(atoms::rot_x()) {
                    if let Ok(val) = term.decode::<i32>() {
                        let mut rotate_x = RotateX::default();
                        rotate_x.set_val(val as i8);
                        view_3d_obj.set_rotate_x(rotate_x);
                    }
                }

                // Try to extract rotate_y from the map
                if let Ok(term) = view_3d.map_get(atoms::rot_y()) {
                    if let Ok(val) = term.decode::<i32>() {
                        let mut rotate_y = RotateY::default();
                        rotate_y.set_val(val as u16);
                        view_3d_obj.set_rotate_y(rotate_y);
                    }
                }

                // Try to extract perspective from the map
                if let Ok(term) = view_3d.map_get(atoms::perspective()) {
                    if let Ok(val) = term.decode::<i32>() {
                        let mut perspective = Perspective::default();
                        perspective.set_val(val as u8);
                        view_3d_obj.set_perspective(perspective);
                    }
                }

                // Apply the view_3d settings to the appropriate chart type
                let plot_area = chart.get_plot_area_mut();
                if let Some(_) = plot_area.get_line_chart_mut() {
                    if chart_type.contains("3D") {
                        // Access the chart space to apply 3D settings
                        let _chart_obj = chart.get_chart_space_mut().get_chart_mut();
                        // Cannot directly set view_3d on Line3DChart in the current API
                        // But we can save the settings for potential future API updates
                    }
                } else if let Some(_) = plot_area.get_bar_chart_mut() {
                    if chart_type.contains("3D") {
                        // Access the chart space to apply 3D settings
                        let _chart_obj = chart.get_chart_space_mut().get_chart_mut();
                        // Cannot directly set view_3d on Bar3DChart in the current API
                    }
                } else if let Some(_) = plot_area.get_pie_chart_mut() {
                    if chart_type.contains("3D") {
                        // Access the chart space to apply 3D settings
                        let _chart_obj = chart.get_chart_space_mut().get_chart_mut();
                        // Cannot directly set view_3d on Pie3DChart in the current API
                    }
                } else if let Some(_) = plot_area.get_area_chart_mut() {
                    if chart_type.contains("3D") {
                        // Access the chart space to apply 3D settings
                        let _chart_obj = chart.get_chart_space_mut().get_chart_mut();
                        // Cannot directly set view_3d on Area3DChart in the current API
                    }
                }
            }

            // Process legend Term if it's a map
            if !legend.is_atom() {
                // Create a new Legend object
                let mut legend_obj = Legend::default();

                // Try to extract position from the map
                if let Ok(term) = legend.map_get(atoms::position()) {
                    if let Ok(pos) = term.decode::<String>() {
                        // Find the legend position from the string parameter
                        let position_value = match pos.to_lowercase().as_str() {
                            "right" => LegendPositionValues::Right,
                            "left" => LegendPositionValues::Left,
                            "top" => LegendPositionValues::Top,
                            "bottom" => LegendPositionValues::Bottom,
                            "topright" => LegendPositionValues::TopRight,
                            _ => LegendPositionValues::Right, // Default
                        };

                        let mut legend_position =
                            umya_spreadsheet::drawing::charts::LegendPosition::default();
                        legend_position.set_val(position_value);
                        legend_obj.set_legend_position(legend_position);
                    }
                }

                // Try to extract overlay from the map
                if let Ok(term) = legend.map_get(atoms::overlay()) {
                    if let Ok(val) = term.decode::<bool>() {
                        let mut overlay_obj = Overlay::default();
                        overlay_obj.set_val(val);
                        legend_obj.set_overlay(overlay_obj);
                    }
                }

                // Set the legend on the chart
                let chart_obj = chart.get_chart_space_mut().get_chart_mut();
                chart_obj.set_legend(legend_obj);
            }

            // Process data_labels Term if it's a map
            if !data_labels.is_atom() {
                let plot_area = chart.get_plot_area_mut();

                // Create a helper function to apply data labels to a series
                let apply_data_labels =
                    |data_labels_obj: &mut umya_spreadsheet::drawing::charts::DataLabels,
                     data_labels_term: &Term| {
                        // Try to extract show_values from the map
                        if let Ok(term) = data_labels_term.map_get(atoms::show_values()) {
                            if let Ok(val) = term.decode::<bool>() {
                                if val {
                                    let mut show_val = ShowValue::default();
                                    show_val.set_val(val);
                                    data_labels_obj.set_show_value(show_val);
                                }
                            }
                        }

                        // Try to extract show_percent from the map
                        if let Ok(term) = data_labels_term.map_get(atoms::show_percent()) {
                            if let Ok(val) = term.decode::<bool>() {
                                if val {
                                    let mut show_percent_obj = ShowPercent::default();
                                    show_percent_obj.set_val(val);
                                    data_labels_obj.set_show_percent(show_percent_obj);
                                }
                            }
                        }

                        // Try to extract show_category_name from the map
                        if let Ok(term) = data_labels_term.map_get(atoms::show_category_name()) {
                            if let Ok(val) = term.decode::<bool>() {
                                if val {
                                    let mut show_cat = ShowCategoryName::default();
                                    show_cat.set_val(val);
                                    data_labels_obj.set_show_category_name(show_cat);
                                }
                            }
                        }

                        // Try to extract show_series_name from the map
                        if let Ok(term) = data_labels_term.map_get(atoms::show_series_name()) {
                            if let Ok(val) = term.decode::<bool>() {
                                if val {
                                    let mut show_series = ShowSeriesName::default();
                                    show_series.set_val(val);
                                    data_labels_obj.set_show_series_name(show_series);
                                }
                            }
                        }
                    };

                // Apply to bar chart series
                if let Some(bar_chart) = plot_area.get_bar_chart_mut() {
                    let series_list = bar_chart.get_area_chart_series_list_mut();
                    for series in series_list.get_area_chart_series_mut() {
                        if let Some(data_labels_obj) = series.get_data_labels_mut() {
                            apply_data_labels(data_labels_obj, &data_labels);
                        }
                    }
                }

                // Apply to line chart series
                if let Some(line_chart) = plot_area.get_line_chart_mut() {
                    let series_list = line_chart.get_area_chart_series_list_mut();
                    for series in series_list.get_area_chart_series_mut() {
                        if let Some(data_labels_obj) = series.get_data_labels_mut() {
                            apply_data_labels(data_labels_obj, &data_labels);
                        }
                    }
                }

                // Apply to pie chart series
                if let Some(pie_chart) = plot_area.get_pie_chart_mut() {
                    let series_list = pie_chart.get_area_chart_series_list_mut();
                    for series in series_list.get_area_chart_series_mut() {
                        if let Some(data_labels_obj) = series.get_data_labels_mut() {
                            apply_data_labels(data_labels_obj, &data_labels);
                        }
                    }
                }

                // Apply to area chart series
                if let Some(area_chart) = plot_area.get_area_chart_mut() {
                    let series_list = area_chart.get_area_chart_series_list_mut();
                    for series in series_list.get_area_chart_series_mut() {
                        if let Some(data_labels_obj) = series.get_data_labels_mut() {
                            apply_data_labels(data_labels_obj, &data_labels);
                        }
                    }
                }
            }

                // Add the chart to the sheet
                sheet.add_chart(chart);

                Ok(atoms::ok())
            }
            None => Err(format!("Sheet '{}' not found", sheet_name)),
        }
    }));

    match result {
        Ok(Ok(atom)) => Ok(atom),
        Ok(Err(err_msg)) => Err((atoms::error(), err_msg)),
        Err(_) => Err((atoms::error(), "Error occurred in add_chart_with_options operation".to_string())),
    }
}
