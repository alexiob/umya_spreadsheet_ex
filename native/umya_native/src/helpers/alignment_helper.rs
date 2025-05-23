use umya_spreadsheet::{HorizontalAlignmentValues, VerticalAlignmentValues};

pub fn parse_horizontal_alignment(alignment: &str) -> HorizontalAlignmentValues {
    match alignment {
        "center" => HorizontalAlignmentValues::Center,
        "left" => HorizontalAlignmentValues::Left,
        "right" => HorizontalAlignmentValues::Right,
        "justify" => HorizontalAlignmentValues::Justify,
        "distributed" => HorizontalAlignmentValues::Distributed,
        "fill" => HorizontalAlignmentValues::Fill,
        "centerContinuous" => HorizontalAlignmentValues::CenterContinuous,
        "general" | _ => HorizontalAlignmentValues::General,
    }
}

pub fn parse_vertical_alignment(alignment: &str) -> VerticalAlignmentValues {
    match alignment {
        "center" => VerticalAlignmentValues::Center,
        "top" => VerticalAlignmentValues::Top,
        "bottom" => VerticalAlignmentValues::Bottom,
        "justify" => VerticalAlignmentValues::Justify,
        "distributed" => VerticalAlignmentValues::Distributed,
        _ => VerticalAlignmentValues::Bottom,
    }
}
