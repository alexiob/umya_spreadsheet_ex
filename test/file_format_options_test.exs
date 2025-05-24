defmodule UmyaSpreadsheetTest.FileFormatOptionsTest do
  use ExUnit.Case, async: true

  alias UmyaSpreadsheet

  setup do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add some test data for file compression tests
    for row <- 1..10 do
      for col <- 1..10 do
        cell = "#{<<64 + col::utf8>>}#{row}"
        value = "Test data for row #{row}, column #{col}"
        :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", cell, value)
      end
    end

    %{spreadsheet: spreadsheet}
  end

  test "write_with_compression creates valid files with different compression levels", %{spreadsheet: spreadsheet} do
    # Test with minimum compression (0)
    no_compress_path = "compression_test_0.xlsx"
    assert :ok = UmyaSpreadsheet.write_with_compression(spreadsheet, no_compress_path, 0)
    assert File.exists?(no_compress_path)

    # Test with medium compression (5)
    medium_compress_path = "compression_test_5.xlsx"
    assert :ok = UmyaSpreadsheet.write_with_compression(spreadsheet, medium_compress_path, 5)
    assert File.exists?(medium_compress_path)

    # Test with maximum compression (9)
    max_compress_path = "compression_test_9.xlsx"
    assert :ok = UmyaSpreadsheet.write_with_compression(spreadsheet, max_compress_path, 9)
    assert File.exists?(max_compress_path)

    # Verify that files with higher compression are smaller (or the same size)
    no_compress_size = File.stat!(no_compress_path).size
    medium_compress_size = File.stat!(medium_compress_path).size
    max_compress_size = File.stat!(max_compress_path).size

    # Files with higher compression should be the same size or smaller
    # In some cases, they might be very close in size depending on the data
    assert medium_compress_size <= no_compress_size + 10 # +10 to allow for small variations
    assert max_compress_size <= medium_compress_size + 10 # +10 to allow for small variations

    # Verify that files can be read back successfully
    {:ok, loaded_spreadsheet} = UmyaSpreadsheet.read(max_compress_path)
    {:ok, cell_value} = UmyaSpreadsheet.get_cell_value(loaded_spreadsheet, "Sheet1", "A1")
    assert "Test data for row 1, column 1" == cell_value

    # Clean up test files
    File.rm(no_compress_path)
    File.rm(medium_compress_path)
    File.rm(max_compress_path)
  end

  test "write_with_encryption_options creates a password protected file", %{spreadsheet: spreadsheet} do
    encrypted_path = "encrypted_test.xlsx"

    # Use the new encryption function with AES256 algorithm
    assert :ok = UmyaSpreadsheet.write_with_encryption_options(
      spreadsheet,
      encrypted_path,
      "test_password",
      "AES256"
    )

    assert File.exists?(encrypted_path)

    # Clean up
    File.rm(encrypted_path)
  end

  test "to_binary_xlsx converts a spreadsheet to binary data", %{spreadsheet: spreadsheet} do
    # Get the binary representation
    binary = UmyaSpreadsheet.to_binary_xlsx(spreadsheet)

    # Binary should be valid data
    assert is_binary(binary)
    assert byte_size(binary) > 0

    # Binary should be valid XLSX data with correct signature
    # First few bytes of a ZIP file (XLSX is a ZIP file) are: [80, 75, 3, 4]
    # (PK\003\004 signature)
    assert <<80, 75, 3, 4, _rest::binary>> = binary

    # Write the binary to a file and verify it can be read back
    temp_path = "binary_test.xlsx"
    :ok = File.write(temp_path, binary)

    # Read it back to verify it's a valid XLSX
    {:ok, loaded_spreadsheet} = UmyaSpreadsheet.read(temp_path)
    assert {:ok, "Test data for row 1, column 1"} = UmyaSpreadsheet.get_cell_value(loaded_spreadsheet, "Sheet1", "A1")

    # Clean up
    File.rm(temp_path)
  end
end
