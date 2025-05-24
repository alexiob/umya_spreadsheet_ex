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
