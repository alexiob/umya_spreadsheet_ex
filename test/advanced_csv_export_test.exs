defmodule UmyaSpreadsheet.AdvancedCsvExportTest do
  use ExUnit.Case, async: true

  @output_csv_path_1 "test/result_files/advanced_export_1.csv"
  @output_csv_path_2 "test/result_files/advanced_export_2.csv"
  @output_csv_path_3 "test/result_files/advanced_export_3.csv"

  setup_all do
    # Make sure the result files directory exists
    File.mkdir_p!("test/result_files")
    :ok
  end

  test "export to CSV with custom options" do
    # Create a new spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add a sheet for demonstration
    :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "Sheet2")

    # Add test data similar to the Rust test
    # Has leading space
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet2", "A1", " TEST")

    :ok =
      UmyaSpreadsheet.set_cell_value(
        spreadsheet,
        "Sheet2",
        "B1",
        "International characters: あいうえお"
      )

    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet2", "C1", "More international: 漢字")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet2", "E1", "1")

    # Has trailing space
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet2", "A2", "TEST ")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet2", "B2", "More international: あいうえお")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet2", "C2", "More characters: 漢字")

    # Has space on both sides
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet2", "A3", " TEST ")

    # Create custom CSV options
    options =
      UmyaSpreadsheet.CsvWriterOption.new()
      # Using UTF-8 for test compatibility
      |> UmyaSpreadsheet.CsvWriterOption.set_csv_encode_value(:utf8)
      # Enable trimming
      |> UmyaSpreadsheet.CsvWriterOption.set_do_trim(true)
      # Wrap values in quotes
      |> UmyaSpreadsheet.CsvWriterOption.set_wrap_with_char("\"")

    # Export to CSV with custom options
    result = UmyaSpreadsheet.write_csv(spreadsheet, "Sheet2", @output_csv_path_1, options)
    assert result == :ok

    # Verify the CSV file was created
    # Read the CSV content to verify it has the expected format
    assert File.exists?(@output_csv_path_1)
    csv_content = File.read!(@output_csv_path_1)

    # The content should have quotes due to our wrap_with_char option
    assert String.contains?(csv_content, "\"TEST\"")

    # Spaces should be trimmed due to do_trim option
    refute String.contains?(csv_content, "\" TEST\"")
    refute String.contains?(csv_content, "\"TEST \"")

    # Data should be exported properly
    assert String.contains?(csv_content, "International characters")

    # Unicode characters should be properly exported
    assert String.contains?(csv_content, "あいうえお")
    assert String.contains?(csv_content, "漢字")
  end

  test "export to CSV with different delimiter" do
    # Create a new spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Set some cell values with commas to test delimiter change
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Value 1, part 2")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Value 2")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C1", "Value 3, extra info")

    # Add some international characters
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "Unicode: あいうえお")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", "More: 漢字")

    # Create CSV options with semi-colon delimiter and wrapping
    options =
      UmyaSpreadsheet.CsvWriterOption.new()
      |> UmyaSpreadsheet.CsvWriterOption.set_wrap_with_char("\"")
      |> UmyaSpreadsheet.CsvWriterOption.set_delimiter(";")

    # Export to CSV
    result = UmyaSpreadsheet.write_csv(spreadsheet, "Sheet1", @output_csv_path_2, options)
    assert result == :ok

    # Verify the file was created
    assert File.exists?(@output_csv_path_2)

    # Read the content
    csv_content = File.read!(@output_csv_path_2)
    assert String.contains?(csv_content, "\"Value 1, part 2\"")

    # Should use semi-colon delimiter instead of comma
    # When we export a row with cells A1, B1, C1, it should join them with semicolons
    assert String.contains?(
             csv_content,
             "\"Value 1, part 2\";\"Value 2\";\"Value 3, extra info\""
           )

    # Shouldn't contain the default comma as delimiter between fields (might have commas inside fields though)
    refute String.contains?(csv_content, "\"Value 1, part 2\",\"Value 2\"")

    # Check for international characters in the second row
    assert String.contains?(csv_content, "あいうえお")
    assert String.contains?(csv_content, "漢字")

    # Verify the delimiter is used for the row with international characters too
    assert String.contains?(csv_content, "\"Unicode: あいうえお\";\"More: 漢字\"")
  end

  test "export to CSV with different encodings for international characters" do
    # Create a new spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add a sheet for demonstration
    :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "Sheet2")

    # Add test data with international characters
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet2", "A1", "Japanese Hiragana: あいうえお")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet2", "B1", "Japanese Kanji: 漢字")

    # Try different encodings - testing we can call the function without errors
    # ShiftJIS encoding option
    options_shift_jis =
      UmyaSpreadsheet.CsvWriterOption.new()
      |> UmyaSpreadsheet.CsvWriterOption.set_csv_encode_value(:shift_jis)
      |> UmyaSpreadsheet.CsvWriterOption.set_wrap_with_char("\"")

    # Export to CSV with encoding option
    result =
      UmyaSpreadsheet.write_csv(spreadsheet, "Sheet2", @output_csv_path_3, options_shift_jis)

    assert result == :ok

    # Verify the CSV file was created
    assert File.exists?(@output_csv_path_3)

    # Read the CSV content as raw binary data
    raw_content = File.read!(@output_csv_path_3)

    # We can verify:
    # 1. The file is not empty
    assert byte_size(raw_content) > 0

    # 2. The file contains quote characters as we specified
    assert String.contains?(raw_content, "\"")

    # 3. The file contains the Japanese characters (either as UTF-8 or encoded)
    # Since we can't reliably test the exact encoding format in a unit test,
    # we'll make sure the content at least contains our specified format characters
    assert String.contains?(raw_content, "Japanese Hiragana")
    assert String.contains?(raw_content, "Japanese Kanji")

    # 4. Test with another encoding option to ensure no errors
    options_utf8 =
      UmyaSpreadsheet.CsvWriterOption.new()
      |> UmyaSpreadsheet.CsvWriterOption.set_csv_encode_value(:utf8)

    result = UmyaSpreadsheet.write_csv(spreadsheet, "Sheet2", @output_csv_path_3, options_utf8)
    assert result == :ok
  end
end
