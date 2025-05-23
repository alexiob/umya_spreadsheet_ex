use rustler::NifStruct;

#[derive(NifStruct, Clone, Debug)]
#[module = "UmyaSpreadsheetEx.CustomStructs.CustomColor"]
pub struct CustomColor {
    pub argb: String,
}

#[derive(NifStruct, Clone, Debug)]
#[module = "UmyaSpreadsheetEx.CustomStructs.CustomFont"]
pub struct CustomFont {
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub color: Option<CustomColor>,
    // Add other font properties here if needed, like size, name, underline, etc.
    // pub size: Option<f64>,
    // pub name: Option<String>,
    // pub underline: Option<String>, // e.g., "single", "double"
}

#[derive(NifStruct, Clone, Debug)]
#[module = "UmyaSpreadsheetEx.CustomStructs.CustomFill"]
pub struct CustomFill {
    pub pattern_type: Option<String>, // e.g., "solid", "gray125", etc.
    pub fg_color: Option<CustomColor>,
    pub bg_color: Option<CustomColor>,
}

// If CustomConditionalFormatValueObject was intended to be a distinct struct, define it here.
// Based on current usage, it seems the fields are directly passed to NIF functions.
// #[derive(NifStruct, Clone, Debug)]
// #[module = "UmyaSpreadsheetEx.CustomStructs.CustomConditionalFormatValueObject"]
// pub struct CustomConditionalFormatValueObject {
//     pub condition_type: String, // Renamed from "type" to avoid keyword clash if it were a field
//     pub val: Option<String>,
// }

// Define SpreadsheetResource if it's a distinct type from UmyaSpreadsheet
// For now, it seems UmyaSpreadsheet is used directly.
// pub type SpreadsheetResource = UmyaSpreadsheet;
