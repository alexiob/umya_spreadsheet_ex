use rustler::{Atom, Decoder, Error as NifError, NifResult, ResourceArc, Term};
use std::collections::HashMap;
use std::sync::Mutex;
use umya_spreadsheet::{
    helper::{coordinate::index_from_coordinate, html::html_to_richtext},
    structs::{Color, Font, RichText, TextElement},
};

use crate::{atoms, UmyaSpreadsheet};

/// Resource wrapper for RichText
pub struct RichTextResource {
    pub rich_text: Mutex<RichText>,
}

/// Resource wrapper for TextElement
pub struct TextElementResource {
    pub text_element: Mutex<TextElement>,
}

/// Create a new RichText object
#[rustler::nif]
pub fn create_rich_text() -> NifResult<ResourceArc<RichTextResource>> {
    let rich_text = RichText::default();
    Ok(ResourceArc::new(RichTextResource {
        rich_text: Mutex::new(rich_text),
    }))
}

/// Create a RichText object from HTML string
#[rustler::nif]
pub fn create_rich_text_from_html(html: String) -> NifResult<ResourceArc<RichTextResource>> {
    match html_to_richtext(&html) {
        Ok(rich_text) => Ok(ResourceArc::new(RichTextResource {
            rich_text: Mutex::new(rich_text),
        })),
        Err(_) => Err(NifError::Term(Box::new((
            atoms::error(),
            "Failed to parse HTML".to_string(),
        )))),
    }
}

/// Create a TextElement with text and optional font properties
#[rustler::nif]
pub fn create_text_element(
    text: String,
    font_props: Term,
) -> NifResult<ResourceArc<TextElementResource>> {
    let mut text_element = TextElement::default();

    // Set the text
    text_element.set_text(text);

    // Parse font properties if provided
    if let Ok(props_map) = <std::collections::HashMap<String, Term>>::decode(font_props) {
        let mut font = Font::default();

        // Parse bold
        if let Some(bold_term) = props_map.get("bold") {
            if let Ok(bold_val) = bool::decode(*bold_term) {
                font.set_bold(bold_val);
            }
        }

        // Parse italic
        if let Some(italic_term) = props_map.get("italic") {
            if let Ok(italic_val) = bool::decode(*italic_term) {
                font.set_italic(italic_val);
            }
        }

        // Parse underline
        if let Some(underline_term) = props_map.get("underline") {
            if let Ok(underline_val) = bool::decode(*underline_term) {
                if underline_val {
                    font.get_font_underline_mut()
                        .set_val(umya_spreadsheet::structs::UnderlineValues::Single);
                }
            }
        }

        // Parse strikethrough
        if let Some(strikethrough_term) = props_map.get("strikethrough") {
            if let Ok(strikethrough_val) = bool::decode(*strikethrough_term) {
                font.set_strikethrough(strikethrough_val);
            }
        }

        // Parse font size
        if let Some(size_term) = props_map.get("size") {
            if let Ok(size_val) = f64::decode(*size_term) {
                font.set_size(size_val);
            }
        }

        // Parse font name
        if let Some(name_term) = props_map.get("name") {
            if let Ok(name_val) = String::decode(*name_term) {
                font.set_name(name_val);
            }
        }

        // Parse color (accepts hex string like "#FF0000" or "red")
        if let Some(color_term) = props_map.get("color") {
            if let Ok(color_val) = String::decode(*color_term) {
                let mut color = Color::default();
                if color_val.starts_with('#') && color_val.len() == 7 {
                    // Hex color
                    color.set_argb(format!("FF{}", &color_val[1..]));
                } else {
                    // Named color - use as RGB hex
                    color.set_argb(color_val);
                }
                font.set_color(color);
            }
        }

        text_element.set_font(font);
    }

    Ok(ResourceArc::new(TextElementResource {
        text_element: Mutex::new(text_element),
    }))
}

/// Add a TextElement to a RichText object
#[rustler::nif]
pub fn add_text_element_to_rich_text(
    rich_text_res: ResourceArc<RichTextResource>,
    text_element_res: ResourceArc<TextElementResource>,
) -> NifResult<Atom> {
    let mut rich_text = rich_text_res.rich_text.lock().unwrap();
    let text_element = text_element_res.text_element.lock().unwrap();

    rich_text.add_rich_text_elements(text_element.clone());
    Ok(atoms::ok())
}

/// Add formatted text directly to a RichText object
#[rustler::nif]
pub fn add_formatted_text_to_rich_text(
    rich_text_res: ResourceArc<RichTextResource>,
    text: String,
    font_props: Term,
) -> NifResult<Atom> {
    let mut rich_text = rich_text_res.rich_text.lock().unwrap();
    let mut text_element = TextElement::default();

    // Set the text
    text_element.set_text(text);

    // Parse font properties if provided
    if let Ok(props_map) = <std::collections::HashMap<String, Term>>::decode(font_props) {
        let mut font = Font::default();

        // Parse bold
        if let Some(bold_term) = props_map.get("bold") {
            if let Ok(bold_val) = bool::decode(*bold_term) {
                font.set_bold(bold_val);
            }
        }

        // Parse italic
        if let Some(italic_term) = props_map.get("italic") {
            if let Ok(italic_val) = bool::decode(*italic_term) {
                font.set_italic(italic_val);
            }
        }

        // Parse underline
        if let Some(underline_term) = props_map.get("underline") {
            if let Ok(underline_val) = bool::decode(*underline_term) {
                if underline_val {
                    font.get_font_underline_mut()
                        .set_val(umya_spreadsheet::structs::UnderlineValues::Single);
                }
            }
        }

        // Parse strikethrough
        if let Some(strikethrough_term) = props_map.get("strikethrough") {
            if let Ok(strikethrough_val) = bool::decode(*strikethrough_term) {
                font.set_strikethrough(strikethrough_val);
            }
        }

        // Parse font size
        if let Some(size_term) = props_map.get("size") {
            if let Ok(size_val) = f64::decode(*size_term) {
                font.set_size(size_val);
            }
        }

        // Parse font name
        if let Some(name_term) = props_map.get("name") {
            if let Ok(name_val) = String::decode(*name_term) {
                font.set_name(name_val);
            }
        }

        // Parse color (accepts hex string like "#FF0000" or "red")
        if let Some(color_term) = props_map.get("color") {
            if let Ok(color_val) = String::decode(*color_term) {
                let mut color = Color::default();
                if color_val.starts_with('#') && color_val.len() == 7 {
                    // Hex color
                    color.set_argb(format!("FF{}", &color_val[1..]));
                } else {
                    // Named color - use as RGB hex
                    color.set_argb(color_val);
                }
                font.set_color(color);
            }
        }

        text_element.set_font(font);
    }

    rich_text.add_rich_text_elements(text_element);
    Ok(atoms::ok())
}

/// Set rich text to a cell
#[rustler::nif]
pub fn set_cell_rich_text(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    coordinate: String,
    rich_text_res: ResourceArc<RichTextResource>,
) -> NifResult<Atom> {
    let mut spreadsheet = resource.spreadsheet.lock().unwrap();
    let rich_text = rich_text_res.rich_text.lock().unwrap();

    match spreadsheet.get_sheet_by_name_mut(&sheet_name) {
        Some(worksheet) => {
            let (col, row, _, _) = index_from_coordinate(&coordinate);

            if let (Some(col_idx), Some(row_idx)) = (col, row) {
                worksheet
                    .get_cell_mut((col_idx, row_idx))
                    .set_rich_text(rich_text.clone());

                Ok(atoms::ok())
            } else {
                Err(NifError::Term(Box::new((
                    atoms::error(),
                    "Invalid coordinate",
                ))))
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Get rich text from a cell
#[rustler::nif]
pub fn get_cell_rich_text(
    resource: ResourceArc<UmyaSpreadsheet>,
    sheet_name: String,
    coordinate: String,
) -> NifResult<ResourceArc<RichTextResource>> {
    let spreadsheet = resource.spreadsheet.lock().unwrap();

    match spreadsheet.get_sheet_by_name(&sheet_name) {
        Some(worksheet) => {
            let (col, row, _, _) = index_from_coordinate(&coordinate);

            if let (Some(col_idx), Some(row_idx)) = (col, row) {
                if let Some(cell) = worksheet.get_cell((col_idx, row_idx)) {
                    let cell_value = cell.get_cell_value();
                    let raw_value = cell_value.get_raw_value();

                    if let Some(rich_text) = raw_value.get_rich_text() {
                        Ok(ResourceArc::new(RichTextResource {
                            rich_text: Mutex::new(rich_text.clone()),
                        }))
                    } else {
                        // Return empty rich text if cell doesn't have rich text
                        Ok(ResourceArc::new(RichTextResource {
                            rich_text: Mutex::new(RichText::default()),
                        }))
                    }
                } else {
                    // Return empty rich text if cell doesn't exist
                    Ok(ResourceArc::new(RichTextResource {
                        rich_text: Mutex::new(RichText::default()),
                    }))
                }
            } else {
                Err(NifError::Term(Box::new((
                    atoms::error(),
                    "Invalid coordinate",
                ))))
            }
        }
        None => Err(NifError::Term(Box::new((
            atoms::error(),
            "Sheet not found".to_string(),
        )))),
    }
}

/// Get plain text from RichText
#[rustler::nif]
pub fn get_rich_text_plain_text(rich_text_res: ResourceArc<RichTextResource>) -> NifResult<String> {
    let rich_text = rich_text_res.rich_text.lock().unwrap();
    Ok(rich_text.get_text().to_string())
}

/// Convert RichText to HTML
#[rustler::nif]
pub fn rich_text_to_html(rich_text_res: ResourceArc<RichTextResource>) -> NifResult<String> {
    let rich_text = rich_text_res.rich_text.lock().unwrap();

    let mut html = String::new();

    for element in rich_text.get_rich_text_elements() {
        let text = element.get_text();

        if let Some(font) = element.get_run_properties() {
            let mut style_parts = Vec::new();

            // Add font family
            if !font.get_name().is_empty() {
                style_parts.push(format!("font-family: {}", font.get_name()));
            }

            // Add font size
            if *font.get_size() > 0.0 {
                style_parts.push(format!("font-size: {}pt", font.get_size()));
            }

            // Add color
            let color = font.get_color();
            if !color.get_argb().is_empty() {
                let argb = color.get_argb();
                if argb.len() >= 8 {
                    let rgb = &argb[2..]; // Remove alpha channel
                    style_parts.push(format!("color: #{}", rgb));
                }
            }

            let style_attr = if !style_parts.is_empty() {
                format!(" style=\"{}\"", style_parts.join("; "))
            } else {
                String::new()
            };

            let mut tag_start = String::new();
            let mut tag_end = String::new();

            // Build nested tags for formatting
            if *font.get_bold() {
                tag_start.push_str("<b>");
                tag_end = "</b>".to_string() + &tag_end;
            }

            if *font.get_italic() {
                tag_start.push_str("<i>");
                tag_end = "</i>".to_string() + &tag_end;
            }

            if *font.get_strikethrough() {
                tag_start.push_str("<s>");
                tag_end = "</s>".to_string() + &tag_end;
            }

            if font.get_font_underline().get_val()
                != &umya_spreadsheet::structs::UnderlineValues::None
            {
                tag_start.push_str("<u>");
                tag_end = "</u>".to_string() + &tag_end;
            }

            if !style_attr.is_empty() {
                tag_start = format!("<span{}>{}", style_attr, tag_start);
                tag_end = tag_end + "</span>";
            }

            html.push_str(&format!("{}{}{}", tag_start, text, tag_end));
        } else {
            html.push_str(text);
        }
    }

    Ok(html)
}

/// Get text elements from RichText
#[rustler::nif]
pub fn get_rich_text_elements(
    rich_text_res: ResourceArc<RichTextResource>,
) -> NifResult<Vec<ResourceArc<TextElementResource>>> {
    let rich_text = rich_text_res.rich_text.lock().unwrap();

    let elements: Vec<ResourceArc<TextElementResource>> = rich_text
        .get_rich_text_elements()
        .iter()
        .map(|element| {
            ResourceArc::new(TextElementResource {
                text_element: Mutex::new(element.clone()),
            })
        })
        .collect();

    Ok(elements)
}

/// Get text from TextElement
#[rustler::nif]
pub fn get_text_element_text(
    text_element_res: ResourceArc<TextElementResource>,
) -> NifResult<String> {
    let text_element = text_element_res.text_element.lock().unwrap();
    Ok(text_element.get_text().to_string())
}

/// Get font properties from TextElement
#[rustler::nif]
pub fn get_text_element_font_properties(
    text_element_res: ResourceArc<TextElementResource>,
) -> NifResult<HashMap<String, String>> {
    let text_element = text_element_res.text_element.lock().unwrap();
    let mut props = HashMap::new();

    // Get font or use default
    if let Some(font) = text_element.get_font() {
        // Get bold property
        props.insert("bold".to_string(), font.get_bold().to_string());

        // Get italic property
        props.insert("italic".to_string(), font.get_italic().to_string());

        // Get strikethrough property
        props.insert(
            "strikethrough".to_string(),
            font.get_strikethrough().to_string(),
        );

        // Get size property
        props.insert("size".to_string(), font.get_size().to_string());

        // Get underline property
        let underline = font.get_font_underline();
        if underline.get_val() != &umya_spreadsheet::structs::UnderlineValues::None {
            props.insert("underline".to_string(), "true".to_string());
        } else {
            props.insert("underline".to_string(), "false".to_string());
        }

        // Get font name
        let name = font.get_name();
        props.insert("name".to_string(), name.to_string());

        // Get color property
        let color = font.get_color();
        let rgb = color.get_argb();
        if !rgb.is_empty() {
            props.insert("color".to_string(), rgb.to_string());
        }
    } else {
        // Default values when no font is set
        props.insert("bold".to_string(), "false".to_string());
        props.insert("italic".to_string(), "false".to_string());
        props.insert("strikethrough".to_string(), "false".to_string());
        props.insert("size".to_string(), "11.0".to_string());
        props.insert("underline".to_string(), "false".to_string());
        props.insert("name".to_string(), "Calibri".to_string());
    }

    Ok(props)
}
