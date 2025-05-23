defmodule UmyaSpreadsheet.LargeStringTest do
  use ExUnit.Case, async: true

  @output_path "test/result_files/large_string_test.xlsx"

  setup do
    # Clean up any previous test files
    File.rm(@output_path)

    %{}
  end

  test "create spreadsheet with large string content" do
    # Create a new spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add a new sheet
    assert :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "LargeDataSheet")

    # Add large amount of cell content
    # We'll add a more reasonable amount for testing purposes (not 5000x30 like in Rust test)
    rows = 200
    cols = 20

    for r <- 1..rows, c <- 1..cols do
      cell_address = "#{<<64 + c::utf8>>}#{r}"
      value = "r#{r}c#{c}"

      assert :ok =
               UmyaSpreadsheet.set_cell_value(spreadsheet, "LargeDataSheet", cell_address, value)
    end

    # Save the spreadsheet
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)

    # Verify file exists and has a reasonable size
    assert File.exists?(@output_path)
    file_size = File.stat!(@output_path).size

    # The file size should be non-zero and reasonably large
    assert file_size > 10_000
  end
end
