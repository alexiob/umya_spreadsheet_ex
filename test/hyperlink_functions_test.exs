defmodule UmyaSpreadsheetTest.HyperlinkFunctionsTest do
  use ExUnit.Case, async: true

  alias UmyaSpreadsheet

  setup do
    # Create a new spreadsheet for each test
    {:ok, spreadsheet} = UmyaSpreadsheet.new()
    %{spreadsheet: spreadsheet}
  end

  describe "hyperlink functions" do
    test "add_hyperlink and get_hyperlink with external URL", %{spreadsheet: spreadsheet} do
      sheet_name = "Sheet1"
      cell_address = "A1"
      url = "https://example.com"
      tooltip = "Visit Example.com"

      # Add a hyperlink to cell A1
      assert :ok =
               UmyaSpreadsheet.add_hyperlink(
                 spreadsheet,
                 sheet_name,
                 cell_address,
                 url,
                 tooltip,
                 false
               )

      # Retrieve the hyperlink
      assert {:ok, hyperlink_info} =
               UmyaSpreadsheet.get_hyperlink(spreadsheet, sheet_name, cell_address)

      assert is_map(hyperlink_info)
      assert hyperlink_info["url"] == url
      assert hyperlink_info["tooltip"] == tooltip
    end

    test "add_hyperlink with minimal parameters", %{spreadsheet: spreadsheet} do
      sheet_name = "Sheet1"
      cell_address = "B2"
      url = "https://google.com"

      # Add a hyperlink with no tooltip and not internal
      assert :ok = UmyaSpreadsheet.add_hyperlink(spreadsheet, sheet_name, cell_address, url)

      # Retrieve the hyperlink
      assert {:ok, hyperlink_info} =
               UmyaSpreadsheet.get_hyperlink(spreadsheet, sheet_name, cell_address)

      assert hyperlink_info["url"] == url
      assert hyperlink_info["tooltip"] == ""
    end

    test "add_hyperlink as internal link", %{spreadsheet: spreadsheet} do
      sheet_name = "Sheet1"
      cell_address = "C3"
      url = "#Sheet2!A1"
      tooltip = "Go to Sheet2 A1"

      # Add an internal hyperlink
      assert :ok =
               UmyaSpreadsheet.add_hyperlink(
                 spreadsheet,
                 sheet_name,
                 cell_address,
                 url,
                 tooltip,
                 true
               )

      # Retrieve the hyperlink
      assert {:ok, hyperlink_info} =
               UmyaSpreadsheet.get_hyperlink(spreadsheet, sheet_name, cell_address)

      assert hyperlink_info["url"] == url
      assert hyperlink_info["tooltip"] == tooltip
      assert hyperlink_info["location"] == cell_address
    end

    test "get_hyperlink returns error for cell without hyperlink", %{spreadsheet: spreadsheet} do
      sheet_name = "Sheet1"
      cell_address = "Z99"

      # Try to get hyperlink from cell that doesn't have one
      assert {:error, _reason} =
               UmyaSpreadsheet.get_hyperlink(spreadsheet, sheet_name, cell_address)
    end

    test "remove_hyperlink", %{spreadsheet: spreadsheet} do
      sheet_name = "Sheet1"
      cell_address = "D4"
      url = "https://test.com"

      # Add a hyperlink first
      :ok = UmyaSpreadsheet.add_hyperlink(spreadsheet, sheet_name, cell_address, url)

      # Verify hyperlink exists
      assert {:ok, _hyperlink_info} =
               UmyaSpreadsheet.get_hyperlink(spreadsheet, sheet_name, cell_address)

      # Remove the hyperlink
      assert :ok = UmyaSpreadsheet.remove_hyperlink(spreadsheet, sheet_name, cell_address)

      # Verify hyperlink is removed
      assert {:error, _reason} =
               UmyaSpreadsheet.get_hyperlink(spreadsheet, sheet_name, cell_address)
    end

    test "has_hyperlink", %{spreadsheet: spreadsheet} do
      sheet_name = "Sheet1"
      cell_address = "E5"
      url = "https://haslink.com"

      # Initially no hyperlink
      assert {:ok, false} = UmyaSpreadsheet.has_hyperlink(spreadsheet, sheet_name, cell_address)

      # Add a hyperlink
      :ok = UmyaSpreadsheet.add_hyperlink(spreadsheet, sheet_name, cell_address, url)

      # Now it should have a hyperlink
      assert {:ok, true} = UmyaSpreadsheet.has_hyperlink(spreadsheet, sheet_name, cell_address)

      # Remove hyperlink
      :ok = UmyaSpreadsheet.remove_hyperlink(spreadsheet, sheet_name, cell_address)

      # Should no longer have hyperlink
      assert {:ok, false} = UmyaSpreadsheet.has_hyperlink(spreadsheet, sheet_name, cell_address)
    end

    test "has_hyperlinks for worksheet", %{spreadsheet: spreadsheet} do
      sheet_name = "Sheet1"

      # Initially no hyperlinks
      assert {:ok, false} = UmyaSpreadsheet.has_hyperlinks(spreadsheet, sheet_name)

      # Add some hyperlinks
      :ok = UmyaSpreadsheet.add_hyperlink(spreadsheet, sheet_name, "F6", "https://first.com")
      :ok = UmyaSpreadsheet.add_hyperlink(spreadsheet, sheet_name, "G7", "https://second.com")

      # Now worksheet should have hyperlinks
      assert {:ok, true} = UmyaSpreadsheet.has_hyperlinks(spreadsheet, sheet_name)

      # Remove one hyperlink
      :ok = UmyaSpreadsheet.remove_hyperlink(spreadsheet, sheet_name, "F6")

      # Should still have hyperlinks
      assert {:ok, true} = UmyaSpreadsheet.has_hyperlinks(spreadsheet, sheet_name)

      # Remove the last hyperlink
      :ok = UmyaSpreadsheet.remove_hyperlink(spreadsheet, sheet_name, "G7")

      # Should no longer have hyperlinks
      assert {:ok, false} = UmyaSpreadsheet.has_hyperlinks(spreadsheet, sheet_name)
    end

    test "get_hyperlinks returns all hyperlinks in worksheet", %{spreadsheet: spreadsheet} do
      sheet_name = "Sheet1"

      # Initially no hyperlinks
      assert {:ok, []} = UmyaSpreadsheet.get_hyperlinks(spreadsheet, sheet_name)

      # Add multiple hyperlinks
      :ok =
        UmyaSpreadsheet.add_hyperlink(
          spreadsheet,
          sheet_name,
          "H8",
          "https://link1.com",
          "Link 1"
        )

      :ok =
        UmyaSpreadsheet.add_hyperlink(
          spreadsheet,
          sheet_name,
          "I9",
          "https://link2.com",
          "Link 2"
        )

      :ok =
        UmyaSpreadsheet.add_hyperlink(
          spreadsheet,
          sheet_name,
          "J10",
          "#Sheet2!B2",
          "Internal Link",
          true
        )

      # Get all hyperlinks
      assert {:ok, hyperlinks} = UmyaSpreadsheet.get_hyperlinks(spreadsheet, sheet_name)
      assert length(hyperlinks) == 3

      # Verify each hyperlink contains the expected keys
      Enum.each(hyperlinks, fn hyperlink ->
        assert Map.has_key?(hyperlink, "cell")
        assert Map.has_key?(hyperlink, "url")
        assert Map.has_key?(hyperlink, "tooltip")
        assert Map.has_key?(hyperlink, "location")
      end)

      # Find specific hyperlinks by cell reference
      h8_link = Enum.find(hyperlinks, fn h -> h["cell"] == "H8" end)
      assert h8_link["url"] == "https://link1.com"
      assert h8_link["tooltip"] == "Link 1"

      i9_link = Enum.find(hyperlinks, fn h -> h["cell"] == "I9" end)
      assert i9_link["url"] == "https://link2.com"
      assert i9_link["tooltip"] == "Link 2"

      j10_link = Enum.find(hyperlinks, fn h -> h["cell"] == "J10" end)
      assert j10_link["url"] == "#Sheet2!B2"
      assert j10_link["tooltip"] == "Internal Link"
      assert j10_link["location"] == "J10"
    end

    test "update_hyperlink", %{spreadsheet: spreadsheet} do
      sheet_name = "Sheet1"
      cell_address = "K11"
      original_url = "https://original.com"
      original_tooltip = "Original tooltip"
      updated_url = "https://updated.com"
      updated_tooltip = "Updated tooltip"

      # Add a hyperlink first
      :ok =
        UmyaSpreadsheet.add_hyperlink(
          spreadsheet,
          sheet_name,
          cell_address,
          original_url,
          original_tooltip
        )

      # Verify original hyperlink
      {:ok, hyperlink_info} = UmyaSpreadsheet.get_hyperlink(spreadsheet, sheet_name, cell_address)
      assert hyperlink_info["url"] == original_url
      assert hyperlink_info["tooltip"] == original_tooltip

      # Update the hyperlink
      assert :ok =
               UmyaSpreadsheet.update_hyperlink(
                 spreadsheet,
                 sheet_name,
                 cell_address,
                 updated_url,
                 updated_tooltip
               )

      # Verify updated hyperlink
      {:ok, updated_info} = UmyaSpreadsheet.get_hyperlink(spreadsheet, sheet_name, cell_address)
      assert updated_info["url"] == updated_url
      assert updated_info["tooltip"] == updated_tooltip
    end

    test "update_hyperlink with internal flag", %{spreadsheet: spreadsheet} do
      sheet_name = "Sheet1"
      cell_address = "L12"
      url = "#Sheet3!C3"
      tooltip = "Internal update test"

      # Add a regular hyperlink first
      :ok =
        UmyaSpreadsheet.add_hyperlink(
          spreadsheet,
          sheet_name,
          cell_address,
          "https://external.com"
        )

      # Update to internal hyperlink
      assert :ok =
               UmyaSpreadsheet.update_hyperlink(
                 spreadsheet,
                 sheet_name,
                 cell_address,
                 url,
                 tooltip,
                 true
               )

      # Verify updated hyperlink
      {:ok, hyperlink_info} = UmyaSpreadsheet.get_hyperlink(spreadsheet, sheet_name, cell_address)
      assert hyperlink_info["url"] == url
      assert hyperlink_info["tooltip"] == tooltip
      assert hyperlink_info["location"] == cell_address
    end

    test "update_hyperlink fails for non-existent hyperlink", %{spreadsheet: spreadsheet} do
      sheet_name = "Sheet1"
      cell_address = "M13"

      # Try to update a hyperlink that doesn't exist
      assert {:error, _reason} =
               UmyaSpreadsheet.update_hyperlink(
                 spreadsheet,
                 sheet_name,
                 cell_address,
                 "https://test.com"
               )
    end

    test "error handling for invalid sheet name", %{spreadsheet: spreadsheet} do
      invalid_sheet = "NonExistentSheet"
      cell_address = "A1"
      url = "https://test.com"

      # All hyperlink functions should return error for invalid sheet
      assert {:error, _reason} =
               UmyaSpreadsheet.add_hyperlink(spreadsheet, invalid_sheet, cell_address, url)

      assert {:error, _reason} =
               UmyaSpreadsheet.get_hyperlink(spreadsheet, invalid_sheet, cell_address)

      assert {:error, _reason} =
               UmyaSpreadsheet.remove_hyperlink(spreadsheet, invalid_sheet, cell_address)

      assert {:error, _reason} =
               UmyaSpreadsheet.has_hyperlink(spreadsheet, invalid_sheet, cell_address)

      assert {:error, _reason} = UmyaSpreadsheet.has_hyperlinks(spreadsheet, invalid_sheet)
      assert {:error, _reason} = UmyaSpreadsheet.get_hyperlinks(spreadsheet, invalid_sheet)

      assert {:error, _reason} =
               UmyaSpreadsheet.update_hyperlink(spreadsheet, invalid_sheet, cell_address, url)
    end

    test "hyperlinks with various URL types", %{spreadsheet: spreadsheet} do
      sheet_name = "Sheet1"

      # Test different URL types
      test_cases = [
        {"N14", "https://secure.example.com", "HTTPS URL"},
        {"O15", "http://example.com", "HTTP URL"},
        {"P16", "ftp://files.example.com", "FTP URL"},
        {"Q17", "mailto:test@example.com", "Email link"},
        {"R18", "file:///C:/Documents/file.pdf", "File path"},
        {"S19", "#Sheet2!A1:B2", "Range reference"}
      ]

      Enum.each(test_cases, fn {cell, url, tooltip} ->
        assert :ok = UmyaSpreadsheet.add_hyperlink(spreadsheet, sheet_name, cell, url, tooltip)

        assert {:ok, hyperlink_info} =
                 UmyaSpreadsheet.get_hyperlink(spreadsheet, sheet_name, cell)

        assert hyperlink_info["url"] == url
        assert hyperlink_info["tooltip"] == tooltip
      end)

      # Verify all hyperlinks are present
      {:ok, all_hyperlinks} = UmyaSpreadsheet.get_hyperlinks(spreadsheet, sheet_name)
      assert length(all_hyperlinks) >= length(test_cases)
    end

    test "hyperlink with empty tooltip", %{spreadsheet: spreadsheet} do
      sheet_name = "Sheet1"
      cell_address = "T20"
      url = "https://notip.com"

      # Add hyperlink with empty tooltip
      assert :ok = UmyaSpreadsheet.add_hyperlink(spreadsheet, sheet_name, cell_address, url, "")

      # Verify hyperlink
      {:ok, hyperlink_info} = UmyaSpreadsheet.get_hyperlink(spreadsheet, sheet_name, cell_address)
      assert hyperlink_info["url"] == url
      assert hyperlink_info["tooltip"] == ""
    end

    test "overwrite existing hyperlink with add_hyperlink", %{spreadsheet: spreadsheet} do
      sheet_name = "Sheet1"
      cell_address = "U21"
      first_url = "https://first.com"
      second_url = "https://second.com"

      # Add first hyperlink
      :ok = UmyaSpreadsheet.add_hyperlink(spreadsheet, sheet_name, cell_address, first_url)

      # Verify first hyperlink
      {:ok, hyperlink_info} = UmyaSpreadsheet.get_hyperlink(spreadsheet, sheet_name, cell_address)
      assert hyperlink_info["url"] == first_url

      # Add second hyperlink to same cell (should overwrite)
      :ok = UmyaSpreadsheet.add_hyperlink(spreadsheet, sheet_name, cell_address, second_url)

      # Verify second hyperlink overwrote the first
      {:ok, updated_info} = UmyaSpreadsheet.get_hyperlink(spreadsheet, sheet_name, cell_address)
      assert updated_info["url"] == second_url
    end
  end

  describe "hyperlink integration with cell values" do
    test "hyperlink on cell with text value", %{spreadsheet: spreadsheet} do
      sheet_name = "Sheet1"
      cell_address = "V22"
      cell_value = "Click here"
      url = "https://clickhere.com"
      tooltip = "This cell has both text and hyperlink"

      # Set cell value first
      :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, cell_address, cell_value)

      # Add hyperlink to the same cell
      :ok = UmyaSpreadsheet.add_hyperlink(spreadsheet, sheet_name, cell_address, url, tooltip)

      # Verify both cell value and hyperlink exist
      {:ok, retrieved_value} =
        UmyaSpreadsheet.get_cell_value(spreadsheet, sheet_name, cell_address)

      assert retrieved_value == cell_value

      {:ok, hyperlink_info} = UmyaSpreadsheet.get_hyperlink(spreadsheet, sheet_name, cell_address)
      assert hyperlink_info["url"] == url
      assert hyperlink_info["tooltip"] == tooltip
    end

    test "removing hyperlink preserves cell value", %{spreadsheet: spreadsheet} do
      sheet_name = "Sheet1"
      cell_address = "W23"
      cell_value = "Important data"
      url = "https://data.com"

      # Set cell value and hyperlink
      :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, sheet_name, cell_address, cell_value)
      :ok = UmyaSpreadsheet.add_hyperlink(spreadsheet, sheet_name, cell_address, url)

      # Remove hyperlink
      :ok = UmyaSpreadsheet.remove_hyperlink(spreadsheet, sheet_name, cell_address)

      # Verify cell value is preserved but hyperlink is gone
      {:ok, retrieved_value} =
        UmyaSpreadsheet.get_cell_value(spreadsheet, sheet_name, cell_address)

      assert retrieved_value == cell_value

      assert {:error, _reason} =
               UmyaSpreadsheet.get_hyperlink(spreadsheet, sheet_name, cell_address)
    end
  end
end
