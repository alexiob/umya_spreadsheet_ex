use umya_spreadsheet::structs::{CsvEncodeValues, CsvWriterOption};

pub fn parse_csv_encoding(encoding: &str) -> CsvEncodeValues {
    match encoding {
        "ShiftJis" => CsvEncodeValues::ShiftJis,
        "Koi8u" => CsvEncodeValues::Koi8u,
        "Koi8r" => CsvEncodeValues::Koi8r,
        "Iso88598i" => CsvEncodeValues::Iso88598i,
        "Gbk" => CsvEncodeValues::Gbk,
        _ => CsvEncodeValues::Utf8,
    }
}

pub fn create_csv_writer_options(
    encoding: &str,
    do_trim: bool,
    wrap_with_char: &str,
) -> CsvWriterOption {
    let mut option = CsvWriterOption::default();

    // Set encoding
    option.set_csv_encode_value(parse_csv_encoding(encoding));

    // Set trimming option
    option.set_do_trim(do_trim);

    // Set wrapping character if provided
    if !wrap_with_char.is_empty() {
        option.set_wrap_with_char(wrap_with_char.to_string());
    }

    option
}
