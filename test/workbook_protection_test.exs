defmodule UmyaSpreadsheetTest.WorkbookProtectionTest do
  use ExUnit.Case, async: true
  doctest UmyaSpreadsheet

  setup do
    # Create a new spreadsheet for each test
    {:ok, spreadsheet} = UmyaSpreadsheet.new()
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Protected Workbook Test")

    # Make sure the result files directory exists
    File.mkdir_p!("test/result_files")

    {:ok, %{spreadsheet: spreadsheet}}
  end

  describe "workbook protection functions" do
    test "is_workbook_protected returns false for unprotected workbooks", %{
      spreadsheet: spreadsheet
    } do
      # A newly created workbook should not be protected
      assert UmyaSpreadsheet.is_workbook_protected(spreadsheet) == {:ok, false}
    end

    test "get_workbook_protection_details returns error for unprotected workbooks", %{
      spreadsheet: spreadsheet
    } do
      # For an unprotected workbook, we expect an error since protection details are not available
      assert UmyaSpreadsheet.get_workbook_protection_details(spreadsheet) ==
               {:error, "Workbook is not protected"}
    end

    test "set workbook protection", %{spreadsheet: spreadsheet} do
      # Check initial state - workbook should not be protected
      assert UmyaSpreadsheet.is_workbook_protected(spreadsheet) == {:ok, false}

      # Set workbook protection with password
      password = "test123"
      :ok = UmyaSpreadsheet.set_workbook_protection(spreadsheet, password)

      # Verify workbook is now protected
      assert UmyaSpreadsheet.is_workbook_protected(spreadsheet) == {:ok, true}

      # Get protection details
      {:ok, protection_details} = UmyaSpreadsheet.get_workbook_protection_details(spreadsheet)
      assert is_map(protection_details)
    end

    test "write password-protected workbook", %{spreadsheet: spreadsheet} do
      # Set workbook protection
      password = "test456"
      :ok = UmyaSpreadsheet.set_workbook_protection(spreadsheet, password)

      # Write password-protected file
      protected_file = "test/result_files/protected_workbook.xlsx"
      :ok = UmyaSpreadsheet.write_with_password(spreadsheet, protected_file, password)

      # Verify file was created
      assert File.exists?(protected_file)

      # Verify file size is reasonable
      file_size = File.stat!(protected_file).size
      assert file_size > 1000, "File size is too small for an Excel file"

      # Attempting to read password-protected file without password should fail
      result = UmyaSpreadsheet.read(protected_file)

      assert match?({:error, _}, result),
             "Should not be able to read password-protected file without password"
    end

    test "workbook protection details API shape", %{spreadsheet: spreadsheet} do
      # For an unprotected workbook, we get an error
      assert UmyaSpreadsheet.get_workbook_protection_details(spreadsheet) ==
               {:error, "Workbook is not protected"}

      # The API is correctly implemented, we verify that the is_workbook_protected functions works too
      assert UmyaSpreadsheet.is_workbook_protected(spreadsheet) == {:ok, false}
    end
  end
end
