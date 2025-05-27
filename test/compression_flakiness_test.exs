defmodule UmyaSpreadsheetTest.CompressionFlakynessTest do
  use ExUnit.Case, async: false

  @tag timeout: 60_000
  test "repeated write_with_compression calls work reliably" do
    path = "test/result_files/high_compression.xlsx"

    # Run this test multiple times to check for flakiness
    for i <- 1..10 do
      # Clean up any existing file
      if File.exists?(path), do: File.rm!(path)

      # Create fresh spreadsheet
      {:ok, spreadsheet} = UmyaSpreadsheet.new()
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Test data #{i}")

      # Call the potentially flaky function
      result = UmyaSpreadsheet.write_with_compression(spreadsheet, path, 9)
      assert result == :ok, "Failed to write on iteration #{i}"
      assert File.exists?(path), "File doesn't exist on iteration #{i}"

      # Verify the file is readable
      {:ok, loaded} = UmyaSpreadsheet.read(path)
      {:ok, value} = UmyaSpreadsheet.get_cell_value(loaded, "Sheet1", "A1")
      assert value == "Test data #{i}", "Wrong value on iteration #{i}"

      # Cleanup
      File.rm!(path)
    end
  end
end
