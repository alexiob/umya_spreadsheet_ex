defmodule UmyaSpreadsheetTest.FileFixTest do
  use ExUnit.Case, async: false

  @moduledoc """
  Tests for verifying fixes to flaky file writing operations.

  This test specifically validates that mutex handling in the Rust code
  properly ensures resources are released between operations.
  """

  test "write_with_compression works reliably across multiple iterations" do
    # Test 10 iterations to verify reliability
    for i <- 1..10 do
      path = "test/result_files/compression_test_#{i}.xlsx"

      # Clean up any existing file
      if File.exists?(path), do: File.rm!(path)

      # Create a new spreadsheet
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add some test data
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Test data #{i}")

      # Test the write operation
      result = UmyaSpreadsheet.write_with_compression(spreadsheet, path, 9)
      assert result == :ok

      # Verify file was created
      assert File.exists?(path)
      assert File.stat!(path).size > 0
    end
  end

  test "write_with_password works reliably across multiple iterations" do
    # Test 5 iterations to verify reliability
    for i <- 1..5 do
      path = "test/result_files/password_test_#{i}.xlsx"

      # Clean up any existing file
      if File.exists?(path), do: File.rm!(path)

      # Create a new spreadsheet
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add some test data
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Password test #{i}")

      # Test the write operation
      result = UmyaSpreadsheet.write_with_password(spreadsheet, path, "test123")
      assert result == :ok

      # Verify file was created
      assert File.exists?(path)
      assert File.stat!(path).size > 0
    end
  end

  test "write_with_encryption_options works reliably across multiple iterations" do
    # Test 5 iterations to verify reliability
    for i <- 1..5 do
      path = "test/result_files/encryption_test_#{i}.xlsx"

      # Clean up any existing file
      if File.exists?(path), do: File.rm!(path)

      # Create a new spreadsheet
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Add some test data
      UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Encryption test #{i}")

      # Test the write operation
      result =
        UmyaSpreadsheet.write_with_encryption_options(spreadsheet, path, "secret", "AES256")

      assert result == :ok

      # Verify file was created
      assert File.exists?(path)
      assert File.stat!(path).size > 0
    end
  end
end
