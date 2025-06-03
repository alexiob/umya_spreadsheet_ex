use rustler::Atom;
use umya_spreadsheet::Color;

pub fn parse_color(color_str: &str) -> Result<String, Atom> {
    // Remove any whitespace
    let color_str = color_str.trim();

    // Handle common color names
    match color_str.to_lowercase().as_str() {
        "red" => Ok(Color::COLOR_RED.to_string()),
        "blue" => Ok(Color::COLOR_BLUE.to_string()),
        "green" => Ok(Color::COLOR_GREEN.to_string()),
        "yellow" => Ok(Color::COLOR_YELLOW.to_string()),
        "black" => Ok(Color::COLOR_BLACK.to_string()),
        "white" => Ok(Color::COLOR_WHITE.to_string()),
        _ => {
            // Handle hex format with alpha (#AARRGGBB)
            if color_str.starts_with("#") && color_str.len() == 9 {
                if let (Ok(a), Ok(r), Ok(g), Ok(b)) = (
                    u8::from_str_radix(&color_str[1..3], 16),
                    u8::from_str_radix(&color_str[3..5], 16),
                    u8::from_str_radix(&color_str[5..7], 16),
                    u8::from_str_radix(&color_str[7..9], 16),
                ) {
                    Ok(format!("{:02X}{:02X}{:02X}{:02X}", a, r, g, b))
                } else {
                    Err(crate::atoms::error())
                }
            }
            // Handle hex format (#RRGGBB)
            else if color_str.starts_with("#") && color_str.len() == 7 {
                // Hex color to ARGB
                if let (Ok(r), Ok(g), Ok(b)) = (
                    u8::from_str_radix(&color_str[1..3], 16),
                    u8::from_str_radix(&color_str[3..5], 16),
                    u8::from_str_radix(&color_str[5..7], 16),
                ) {
                    Ok(format!("FF{:02X}{:02X}{:02X}", r, g, b))
                } else {
                    Err(crate::atoms::error())
                }
            }
            // Handle hex format without # (RRGGBB)
            else if color_str.len() == 6 && color_str.chars().all(|c| c.is_ascii_hexdigit()) {
                // Handle color without # prefix (e.g., RRGGBB)
                if let (Ok(r), Ok(g), Ok(b)) = (
                    u8::from_str_radix(&color_str[0..2], 16),
                    u8::from_str_radix(&color_str[2..4], 16),
                    u8::from_str_radix(&color_str[4..6], 16),
                ) {
                    Ok(format!("FF{:02X}{:02X}{:02X}", r, g, b))
                } else {
                    Err(crate::atoms::error())
                }
            }
            // Handle hex format without # (AARRGGBB)
            else if color_str.len() == 8 && color_str.chars().all(|c| c.is_ascii_hexdigit()) {
                if let (Ok(a), Ok(r), Ok(g), Ok(b)) = (
                    u8::from_str_radix(&color_str[0..2], 16),
                    u8::from_str_radix(&color_str[2..4], 16),
                    u8::from_str_radix(&color_str[4..6], 16),
                    u8::from_str_radix(&color_str[6..8], 16),
                ) {
                    Ok(format!("{:02X}{:02X}{:02X}{:02X}", a, r, g, b))
                } else {
                    Err(crate::atoms::error())
                }
            } else {
                Err(crate::atoms::error())
            }
        }
    }
}

pub fn create_color_object(color_str: &str) -> Result<Color, Atom> {
    let argb = parse_color(color_str)?;
    let mut color_obj = Color::default();
    color_obj.set_argb(argb);
    Ok(color_obj)
}
