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

    # We can verify encryption is applied by:

    # 1. Attempting to read the file without a password should fail
    # This verifies the password protection is working
    result = UmyaSpreadsheet.read(@output_password_file_path)

    assert match?({:error, _}, result),
           "Expected error when opening password-protected file without password"

    # 2. Check file properties that indicate encryption
    # The encrypted file should be non-empty and different from the original
    input_file_size = File.stat!(@input_file_path).size
    output_file_size = File.stat!(@output_password_file_path).size
    assert output_file_size > 0, "Output file should not be empty"

    assert output_file_size != input_file_size,
           "Output file size should differ due to encryption"

    # 3. The relation between sizes is unpredictable but generally within a reasonable range
    # (Encrypted files are typically within 50% of the original size)
    assert_in_delta output_file_size, input_file_size, input_file_size * 0.5
  end
end
