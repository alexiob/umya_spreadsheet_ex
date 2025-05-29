defmodule UmyaSpreadsheet.WorkbookProtectionFunctions do
  @moduledoc """
  Functions for managing workbook protection settings and retrieving protection status.

  Excel workbooks can be protected to prevent structural changes, window modifications, or revision tracking.
  This module provides functions to:
  * Check if workbook protection is enabled
  * Retrieve protection settings and status
  * Enable/disable various protection features
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaNative

  @doc """
  Checks if the workbook has protection enabled.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct

  ## Returns

  - `{:ok, is_protected}` where is_protected is a boolean indicating if protection is enabled
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, protected} = UmyaSpreadsheet.WorkbookProtectionFunctions.is_workbook_protected(spreadsheet)
      # protected = true (if workbook is protected)
  """
  def is_workbook_protected(%Spreadsheet{reference: ref}) do
    UmyaNative.is_workbook_protected(ref)
  end

  @doc """
  Gets detailed information about workbook protection settings.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct

  ## Returns

  - `{:ok, protection_details}` where protection_details is a map containing protection settings:
    - `"lock_structure"` - "true" if structure changes are prevented
    - `"lock_windows"` - "true" if window changes are prevented
    - `"lock_revision"` - "true" if revision tracking is enabled
  - `{:error, reason}` if the workbook is not protected or on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("protected.xlsx")
      {:ok, details} = UmyaSpreadsheet.WorkbookProtectionFunctions.get_workbook_protection_details(spreadsheet)
      # details = %{
      #   "lock_structure" => "true",
      #   "lock_windows" => "true",
      #   "lock_revision" => "false"
      # }
  """
  def get_workbook_protection_details(%Spreadsheet{reference: ref}) do
    UmyaNative.get_workbook_protection_details(ref)
  end
end
