defmodule UmyaSpreadsheetTest.WriteWithCompressionTest do
  use ExUnit.Case, async: false

  alias UmyaSpreadsheet

  test "write_with_compression should work reliably for multiple iterations" do
    # Try 10 iterations to test reliability
    iterations = 10

    results =
      for i <- 1..iterations do
        # Define a unique path for this iteration
        path = "test/result_files/compression_test_#{i}.xlsx"

        # Remove the file if it already exists
        if File.exists?(path), do: File.rm!(path)

        # Create a new spreadsheet
        {:ok, spreadsheet} = UmyaSpreadsheet.new()

        # Add some data
        UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Test data #{i}")

        # Try to write with compression level 9
        result = UmyaSpreadsheet.write_with_compression(spreadsheet, path, 9)

        # Print progress
        IO.puts("Iteration #{i}/#{iterations}, Result: #{inspect(result)}")

        # Check if file was created
        file_exists = File.exists?(path)
        IO.puts("  File exists: #{file_exists}")

        if file_exists do
          file_size = File.stat!(path).size
          IO.puts("  File size: #{file_size} bytes")
        end

        # Return success or failure for this iteration
        if result == :ok && file_exists, do: :ok, else: :error
      end

    # Assert all iterations succeeded
    assert Enum.all?(results, &(&1 == :ok)), "Some iterations failed"
  end
end
