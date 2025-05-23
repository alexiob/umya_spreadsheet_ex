use umya_spreadsheet::{Color, Fill, Style};

pub fn parse_color(color_str: &str) -> Color {
    let mut color = Color::default();
    if color_str.starts_with('#') && color_str.len() == 7 {
        color.set_argb(&format!("FF{}", &color_str[1..]));
    } else {
        color.set_argb(color_str);
    }
    color
}

pub fn create_fill_style(color_str: &str) -> Option<Style> {
    let mut style = Style::default();
    let mut fill = Fill::default();
    let pattern = {
        let mut pf = umya_spreadsheet::PatternFill::default();
        if color_str.starts_with('#') && color_str.len() == 7 {
            let mut color = Color::default();
            color.set_argb(&format!("FF{}", &color_str[1..]));
            pf.set_foreground_color(color);
        } else {
            let mut color = Color::default();
            color.set_argb(color_str);
            pf.set_foreground_color(color);
        }
        pf.set_pattern_type(umya_spreadsheet::PatternValues::Solid);
        pf
    };
    fill.set_pattern_fill(pattern);
    style.set_fill(fill);
    Some(style)
}
