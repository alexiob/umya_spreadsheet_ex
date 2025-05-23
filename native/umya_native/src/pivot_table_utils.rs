use crate::UmyaSpreadsheet;
use rustler::ResourceArc;

#[allow(dead_code)]
pub fn has_pivot_tables(sheet_name: &str, spreadsheet: &ResourceArc<UmyaSpreadsheet>) -> bool {
    let guard = spreadsheet.spreadsheet.lock().unwrap();
    if let Some(sheet) = guard.get_sheet_by_name(sheet_name) {
        let pivot_tables = sheet.get_pivot_tables();
        let _count = pivot_tables.iter().count();

        for (_i, pt) in pivot_tables.iter().enumerate() {
            let def = pt.get_pivot_table_definition();
            let has_name = !def.get_name().is_empty();
            let has_cache = *def.get_cache_id() > 0;
            let has_loc = !def.get_location().get_reference().is_empty();
            let has_fields = !def.get_pivot_fields().get_list().is_empty();

            if has_name && has_cache && has_loc && has_fields {
                return true;
            }
        }
    }
    false
}
