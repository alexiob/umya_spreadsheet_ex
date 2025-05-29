use rustler::{Encoder, Env, ResourceArc, Term};
use std::collections::HashMap;
use umya_spreadsheet::helper::coordinate::CellCoordinates;
use umya_spreadsheet::{Coordinate, Table, TableColumn, TableStyleInfo, TotalsRowFunctionValues};

use crate::atoms;
use crate::UmyaSpreadsheet;

// ============================================================================
// TABLE MANAGEMENT FUNCTIONS
// ============================================================================

/// Add a table to a worksheet
#[rustler::nif]
pub fn add_table(
    env: Env,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    table_name: String,
    display_name: String,
    start_cell: String,
    end_cell: String,
    columns: Vec<String>,
    has_totals_row: Option<bool>,
) -> Term {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let mut table = Table::default();

            // Set basic table properties
            table.set_name(&table_name);
            table.set_display_name(&display_name);

            // Set table area using coordinate parsing
            match parse_cell_coordinates(&start_cell, &end_cell) {
                Ok((start_coord, end_coord)) => {
                    table.set_area((start_coord, end_coord));
                }
                Err(_) => {
                    return (atoms::error(), "Invalid cell coordinates".to_string()).encode(env);
                }
            }

            // Add columns to the table
            for column_name in columns {
                let column = TableColumn::new(&column_name);
                table.add_column(column);
            }

            // Set totals row if specified
            if let Some(has_totals) = has_totals_row {
                table.set_totals_row_shown(has_totals);
                if has_totals {
                    table.set_totals_row_count(1);
                }
            }

            sheet.add_table(table);
            (atoms::ok(), atoms::ok()).encode(env)
        }
        None => (atoms::error(), "Sheet not found".to_string()).encode(env),
    }
}

/// Get all tables from a worksheet
#[rustler::nif]
pub fn get_tables(env: Env, resource: ResourceArc<UmyaSpreadsheet>, sheet_name: String) -> Term {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            let mut tables_info = Vec::new();

            for table in sheet.get_tables() {
                let mut table_map = HashMap::new();

                table_map.insert("name".to_string(), table.get_name().encode(env));
                table_map.insert(
                    "display_name".to_string(),
                    table.get_display_name().encode(env),
                );

                // Convert area coordinates to cell references
                let (start_coord, end_coord) = table.get_area();
                let start_cell = coordinate_to_cell_reference(start_coord);
                let end_cell = coordinate_to_cell_reference(end_coord);
                table_map.insert("start_cell".to_string(), start_cell.encode(env));
                table_map.insert("end_cell".to_string(), end_cell.encode(env));

                // Get column information
                let columns: Vec<String> = table
                    .get_columns()
                    .iter()
                    .map(|col| col.get_name().to_string())
                    .collect();
                table_map.insert("columns".to_string(), columns.encode(env));

                table_map.insert(
                    "has_totals_row".to_string(),
                    table.get_totals_row_shown().encode(env),
                );

                // Get style info if available
                if let Some(style_info) = table.get_style_info() {
                    let mut style_map = HashMap::new();
                    style_map.insert("name".to_string(), style_info.get_name().encode(env));
                    style_map.insert(
                        "show_first_col".to_string(),
                        style_info.is_show_first_col().encode(env),
                    );
                    style_map.insert(
                        "show_last_col".to_string(),
                        style_info.is_show_last_col().encode(env),
                    );
                    style_map.insert(
                        "show_row_stripes".to_string(),
                        style_info.is_show_row_stripes().encode(env),
                    );
                    style_map.insert(
                        "show_col_stripes".to_string(),
                        style_info.is_show_col_stripes().encode(env),
                    );
                    table_map.insert("style_info".to_string(), style_map.encode(env));
                }

                tables_info.push(table_map);
            }

            (atoms::ok(), tables_info).encode(env)
        }
        None => (atoms::error(), "Sheet not found".to_string()).encode(env),
    }
}

/// Remove a table from a worksheet by name
#[rustler::nif]
pub fn remove_table(
    env: Env,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    table_name: String,
) -> Term {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let tables = sheet.get_tables_mut();
            let original_count = tables.len();

            // Remove table with matching name
            tables.retain(|table| table.get_name() != table_name);

            if tables.len() < original_count {
                (atoms::ok(), atoms::ok()).encode(env)
            } else {
                (atoms::error(), "Table not found".to_string()).encode(env)
            }
        }
        None => (atoms::error(), "Sheet not found".to_string()).encode(env),
    }
}

/// Check if a sheet has any tables
#[rustler::nif]
pub fn has_tables(env: Env, resource: ResourceArc<UmyaSpreadsheet>, sheet_name: String) -> Term {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => (atoms::ok(), !sheet.get_tables().is_empty()).encode(env),
        None => (atoms::error(), "Sheet not found".to_string()).encode(env),
    }
}

/// Count the number of tables in a worksheet
#[rustler::nif]
pub fn count_tables(env: Env, resource: ResourceArc<UmyaSpreadsheet>, sheet_name: String) -> Term {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => (atoms::ok(), sheet.get_tables().len()).encode(env),
        None => (atoms::error(), "Sheet not found".to_string()).encode(env),
    }
}

// ============================================================================
// TABLE STYLING FUNCTIONS
// ============================================================================

/// Set table style information
#[rustler::nif]
pub fn set_table_style(
    env: Env,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    table_name: String,
    style_name: String,
    show_first_col: bool,
    show_last_col: bool,
    show_row_stripes: bool,
    show_col_stripes: bool,
) -> Term {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let tables = sheet.get_tables_mut();

            // Find the table to modify
            if let Some(table) = tables.iter_mut().find(|t| t.get_name() == table_name) {
                let style_info = TableStyleInfo::new(
                    &style_name,
                    show_first_col,
                    show_last_col,
                    show_row_stripes,
                    show_col_stripes,
                );

                table.set_style_info(Some(style_info));
                (atoms::ok(), atoms::ok()).encode(env)
            } else {
                (atoms::error(), "Table not found".to_string()).encode(env)
            }
        }
        None => (atoms::error(), "Sheet not found".to_string()).encode(env),
    }
}

/// Remove table style information
#[rustler::nif]
pub fn remove_table_style(
    env: Env,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    table_name: String,
) -> Term {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let tables = sheet.get_tables_mut();

            // Find the table to modify
            if let Some(table) = tables.iter_mut().find(|t| t.get_name() == table_name) {
                table.set_style_info(None);
                (atoms::ok(), atoms::ok()).encode(env)
            } else {
                (atoms::error(), "Table not found".to_string()).encode(env)
            }
        }
        None => (atoms::error(), "Sheet not found".to_string()).encode(env),
    }
}

// ============================================================================
// TABLE COLUMN MANAGEMENT FUNCTIONS
// ============================================================================

/// Add a column to an existing table
#[rustler::nif]
pub fn add_table_column(
    env: Env,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    table_name: String,
    column_name: String,
    totals_row_function: Option<String>,
    totals_row_label: Option<String>,
) -> Term {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let tables = sheet.get_tables_mut();

            // Find the table to modify
            if let Some(table) = tables.iter_mut().find(|t| t.get_name() == table_name) {
                let mut column = TableColumn::new(&column_name);

                // Set totals row function if provided
                if let Some(function_str) = totals_row_function {
                    if let Ok(function) = parse_totals_row_function(&function_str) {
                        column.set_totals_row_function(function);
                    }
                }

                // Set totals row label if provided
                if let Some(label) = totals_row_label {
                    column.set_totals_row_label(&label);
                }

                table.add_column(column);
                (atoms::ok(), atoms::ok()).encode(env)
            } else {
                (atoms::error(), "Table not found".to_string()).encode(env)
            }
        }
        None => (atoms::error(), "Sheet not found".to_string()).encode(env),
    }
}

/// Modify an existing table column
#[rustler::nif]
pub fn modify_table_column(
    env: Env,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    table_name: String,
    old_column_name: String,
    new_column_name: Option<String>,
    totals_row_function: Option<String>,
    totals_row_label: Option<String>,
) -> Term {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let tables = sheet.get_tables_mut();

            // Find the table to modify
            if let Some(table) = tables.iter_mut().find(|t| t.get_name() == table_name) {
                // We need to access the columns directly but TableColumn doesn't seem to have
                // a mutable iterator in the public API, so we'll need to work around this
                // For now, let's create a new implementation that recreates the table with modified columns

                // Get current table data
                let current_name = table.get_name().to_string();
                let current_display_name = table.get_display_name().to_string();
                let current_area = table.get_area().clone();
                let current_totals_shown = *table.get_totals_row_shown();
                let current_style = table.get_style_info().cloned();

                // Create new columns list with modifications
                let mut new_columns = Vec::new();
                let mut found_column = false;

                for existing_col in table.get_columns() {
                    if existing_col.get_name() == old_column_name {
                        found_column = true;
                        let mut new_col =
                            TableColumn::new(new_column_name.as_ref().unwrap_or(&old_column_name));

                        // Copy existing settings
                        if let Some(existing_label) = existing_col.get_totals_row_label() {
                            new_col.set_totals_row_label(existing_label);
                        }
                        new_col.set_totals_row_function(
                            existing_col.get_totals_row_function().clone(),
                        );

                        // Apply new settings if provided
                        if let Some(function_str) = &totals_row_function {
                            if let Ok(function) = parse_totals_row_function(function_str) {
                                new_col.set_totals_row_function(function);
                            }
                        }

                        if let Some(label) = &totals_row_label {
                            new_col.set_totals_row_label(label);
                        }

                        new_columns.push(new_col);
                    } else {
                        // Create a copy of the existing column
                        let mut new_col = TableColumn::new(existing_col.get_name());
                        if let Some(existing_label) = existing_col.get_totals_row_label() {
                            new_col.set_totals_row_label(existing_label);
                        }
                        new_col.set_totals_row_function(
                            existing_col.get_totals_row_function().clone(),
                        );
                        new_columns.push(new_col);
                    }
                }

                if !found_column {
                    return (atoms::error(), "Column not found".to_string()).encode(env);
                }

                // Create a new table with updated columns
                let mut new_table = Table::default();
                new_table.set_name(&current_name);
                new_table.set_display_name(&current_display_name);

                // Convert Coordinate to CellCoordinates for set_area
                let start_cell_coord = CellCoordinates::from((
                    *current_area.0.get_col_num(),
                    *current_area.0.get_row_num(),
                ));
                let end_cell_coord = CellCoordinates::from((
                    *current_area.1.get_col_num(),
                    *current_area.1.get_row_num(),
                ));
                new_table.set_area((start_cell_coord, end_cell_coord));

                new_table.set_totals_row_shown(current_totals_shown);

                for col in new_columns {
                    new_table.add_column(col);
                }

                if let Some(style) = current_style {
                    new_table.set_style_info(Some(style));
                }

                // Replace the old table
                *table = new_table;

                (atoms::ok(), atoms::ok()).encode(env)
            } else {
                (atoms::error(), "Table not found".to_string()).encode(env)
            }
        }
        None => (atoms::error(), "Sheet not found".to_string()).encode(env),
    }
}

/// Set totals row visibility for a table
#[rustler::nif]
pub fn set_table_totals_row(
    env: Env,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    table_name: String,
    show_totals_row: bool,
) -> Term {
    let mut guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name_mut(&sheet_name) {
        Some(sheet) => {
            let tables = sheet.get_tables_mut();

            // Find the table to modify
            if let Some(table) = tables.iter_mut().find(|t| t.get_name() == table_name) {
                table.set_totals_row_shown(show_totals_row);
                if show_totals_row {
                    table.set_totals_row_count(1);
                } else {
                    table.set_totals_row_count(0);
                }
                (atoms::ok(), atoms::ok()).encode(env)
            } else {
                (atoms::error(), "Table not found".to_string()).encode(env)
            }
        }
        None => (atoms::error(), "Sheet not found".to_string()).encode(env),
    }
}

// ============================================================================
// TABLE GETTER FUNCTIONS
// ============================================================================

/// Get a specific table by name from a worksheet
#[rustler::nif]
pub fn get_table(
    env: Env,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    table_name: String,
) -> Term {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            let tables = sheet.get_tables();

            if let Some(table) = tables.iter().find(|t| t.get_name() == table_name) {
                let mut table_map = HashMap::new();

                table_map.insert("name".to_string(), table.get_name().to_string());
                table_map.insert(
                    "display_name".to_string(),
                    table.get_display_name().to_string(),
                );

                // Convert area coordinates to cell references
                let (start_coord, end_coord) = table.get_area();
                let start_cell = coordinate_to_cell_reference(start_coord);
                let end_cell = coordinate_to_cell_reference(end_coord);
                table_map.insert("start_cell".to_string(), start_cell);
                table_map.insert("end_cell".to_string(), end_cell);

                // Get column information
                let columns: Vec<String> = table
                    .get_columns()
                    .iter()
                    .map(|col| col.get_name().to_string())
                    .collect();
                table_map.insert("columns".to_string(), format!("{:?}", columns));

                table_map.insert(
                    "has_totals_row".to_string(),
                    table.get_totals_row_shown().to_string(),
                );

                // Get style info if available
                if let Some(style_info) = table.get_style_info() {
                    let mut style_map = HashMap::new();
                    style_map.insert("name".to_string(), style_info.get_name().to_string());
                    style_map.insert(
                        "show_first_col".to_string(),
                        style_info.is_show_first_col().to_string(),
                    );
                    style_map.insert(
                        "show_last_col".to_string(),
                        style_info.is_show_last_col().to_string(),
                    );
                    style_map.insert(
                        "show_row_stripes".to_string(),
                        style_info.is_show_row_stripes().to_string(),
                    );
                    style_map.insert(
                        "show_col_stripes".to_string(),
                        style_info.is_show_col_stripes().to_string(),
                    );
                    table_map.insert("style_info".to_string(), format!("{:?}", style_map));
                }

                (atoms::ok(), table_map).encode(env)
            } else {
                (atoms::error(), "Table not found".to_string()).encode(env)
            }
        }
        None => (atoms::error(), "Sheet not found".to_string()).encode(env),
    }
}

/// Get table style information for a specific table
#[rustler::nif]
pub fn get_table_style(
    env: Env,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    table_name: String,
) -> Term {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            let tables = sheet.get_tables();

            if let Some(table) = tables.iter().find(|t| t.get_name() == table_name) {
                if let Some(style_info) = table.get_style_info() {
                    let mut style_map = HashMap::new();

                    style_map.insert("name".to_string(), style_info.get_name().to_string());
                    style_map.insert(
                        "show_first_column".to_string(),
                        style_info.is_show_first_col().to_string(),
                    );
                    style_map.insert(
                        "show_last_column".to_string(),
                        style_info.is_show_last_col().to_string(),
                    );
                    style_map.insert(
                        "show_row_stripes".to_string(),
                        style_info.is_show_row_stripes().to_string(),
                    );
                    style_map.insert(
                        "show_column_stripes".to_string(),
                        style_info.is_show_col_stripes().to_string(),
                    );

                    (atoms::ok(), style_map).encode(env)
                } else {
                    (atoms::ok(), HashMap::<String, String>::new()).encode(env)
                }
            } else {
                (atoms::error(), "Table not found".to_string()).encode(env)
            }
        }
        None => (atoms::error(), "Sheet not found".to_string()).encode(env),
    }
}

/// Get table columns information for a specific table
#[rustler::nif]
pub fn get_table_columns(
    env: Env,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    table_name: String,
) -> Term {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            let tables = sheet.get_tables();

            if let Some(table) = tables.iter().find(|t| t.get_name() == table_name) {
                let columns = table.get_columns();
                let mut columns_info = Vec::new();

                for column in columns {
                    let mut column_map = HashMap::new();
                    column_map.insert("name".to_string(), column.get_name().to_string());

                    if let Some(label) = column.get_totals_row_label() {
                        column_map.insert("totals_row_label".to_string(), label.to_string());
                    }

                    let function = column.get_totals_row_function();
                    let function_str = match function {
                        TotalsRowFunctionValues::Sum => "sum",
                        TotalsRowFunctionValues::Count => "count",
                        TotalsRowFunctionValues::Average => "average",
                        TotalsRowFunctionValues::Maximum => "max",
                        TotalsRowFunctionValues::Minimum => "min",
                        TotalsRowFunctionValues::StandardDeviation => "stddev",
                        TotalsRowFunctionValues::Variance => "var",
                        TotalsRowFunctionValues::CountNumbers => "countnums",
                        TotalsRowFunctionValues::Custom => "custom",
                        TotalsRowFunctionValues::None => "none",
                    };
                    column_map.insert("totals_row_function".to_string(), function_str.to_string());

                    columns_info.push(column_map);
                }

                (atoms::ok(), columns_info).encode(env)
            } else {
                (atoms::error(), "Table not found".to_string()).encode(env)
            }
        }
        None => (atoms::error(), "Sheet not found".to_string()).encode(env),
    }
}

/// Get totals row visibility status for a specific table
#[rustler::nif]
pub fn get_table_totals_row(
    env: Env,
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    table_name: String,
) -> Term {
    let guard = resource.spreadsheet.lock().unwrap();

    match guard.get_sheet_by_name(&sheet_name) {
        Some(sheet) => {
            let tables = sheet.get_tables();

            if let Some(table) = tables.iter().find(|t| t.get_name() == table_name) {
                let totals_row_shown = *table.get_totals_row_shown();
                (atoms::ok(), totals_row_shown).encode(env)
            } else {
                (atoms::error(), "Table not found".to_string()).encode(env)
            }
        }
        None => (atoms::error(), "Sheet not found".to_string()).encode(env),
    }
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

fn parse_cell_coordinates(
    start_cell: &str,
    end_cell: &str,
) -> Result<(CellCoordinates, CellCoordinates), String> {
    // Parse cell references like "A1" into coordinates
    let (start_col, start_row, _, _) =
        umya_spreadsheet::helper::coordinate::index_from_coordinate(start_cell);
    let (end_col, end_row, _, _) =
        umya_spreadsheet::helper::coordinate::index_from_coordinate(end_cell);

    let start_coord = CellCoordinates::from((
        start_col.ok_or("Invalid start cell column")?,
        start_row.ok_or("Invalid start cell row")?,
    ));
    let end_coord = CellCoordinates::from((
        end_col.ok_or("Invalid end cell column")?,
        end_row.ok_or("Invalid end cell row")?,
    ));

    Ok((start_coord, end_coord))
}

fn coordinate_to_cell_reference(coord: &Coordinate) -> String {
    umya_spreadsheet::helper::coordinate::coordinate_from_index(
        coord.get_col_num(),
        coord.get_row_num(),
    )
}

fn parse_totals_row_function(function_str: &str) -> Result<TotalsRowFunctionValues, String> {
    match function_str.to_lowercase().as_str() {
        "sum" => Ok(TotalsRowFunctionValues::Sum),
        "count" => Ok(TotalsRowFunctionValues::Count),
        "average" => Ok(TotalsRowFunctionValues::Average),
        "maximum" | "max" => Ok(TotalsRowFunctionValues::Maximum),
        "minimum" | "min" => Ok(TotalsRowFunctionValues::Minimum),
        "standarddeviation" | "stddev" => Ok(TotalsRowFunctionValues::StandardDeviation),
        "variance" | "var" => Ok(TotalsRowFunctionValues::Variance),
        "countnumbers" | "countnums" => Ok(TotalsRowFunctionValues::CountNumbers),
        "custom" => Ok(TotalsRowFunctionValues::Custom),
        "none" => Ok(TotalsRowFunctionValues::None),
        _ => Err(format!("Unknown totals row function: {}", function_str)),
    }
}
