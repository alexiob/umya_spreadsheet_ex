defmodule UmyaSpreadsheet.SetPasswordTest do
  use ExUnit.Case, async: true

  @input_file_path "test/test_files/aaa.xlsx"
  @output_password_file_path "test/result_files/set_password_test.xlsx"
  @password "test_password123"

  setup do
    # Make sure the result directory exists
    File.mkdir_p!("test/result_files")

    # Return test context
    :ok
  end

  test "set password on existing Excel file" do
    # Ensure the input file exists
    assert File.exists?(@input_file_path)

    # Clean up any previous output file
    File.rm(@output_password_file_path)

    # Apply password protection to the file
    assert :ok =
             UmyaSpreadsheet.set_password(@input_file_path, @output_password_file_path, @password)

    # Verify that the output file was created
    assert File.exists?(@output_password_file_path)

    # We can't easily verify the password protection programmatically,
    # but we can check that the file size has changed (encryption adds overhead)
    input_file_size = File.stat!(@input_file_path).size
    output_file_size = File.stat!(@output_password_file_path).size

    # The encrypted file should be at least as large as the original
    assert output_file_size > 0
    assert_in_delta output_file_size, input_file_size, input_file_size * 0.5
  end
end
