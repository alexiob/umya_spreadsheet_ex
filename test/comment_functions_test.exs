defmodule UmyaSpreadsheetTest.CommentFunctionsTest do
  use ExUnit.Case

  alias UmyaSpreadsheet

  setup do
    # Create a new spreadsheet for each test
    {:ok, spreadsheet} = UmyaSpreadsheet.new()
    %{spreadsheet: spreadsheet}
  end

  describe "comment functions" do
    test "add_comment and get_comment", %{spreadsheet: spreadsheet} do
      sheet_name = "Sheet1"
      cell_address = "A1"
      comment_text = "This is a test comment"
      author = "Test Author"

      # Add a comment to cell A1
      assert :ok = UmyaSpreadsheet.add_comment(spreadsheet, sheet_name, cell_address, comment_text, author)

      # Retrieve the comment
      assert {:ok, retrieved_text, retrieved_author} = UmyaSpreadsheet.get_comment(spreadsheet, sheet_name, cell_address)
      assert retrieved_text == comment_text
      assert retrieved_author == author
    end

    test "update_comment", %{spreadsheet: spreadsheet} do
      sheet_name = "Sheet1"
      cell_address = "B2"
      comment_text = "Original comment"
      author = "Original Author"
      updated_text = "Updated comment"
      updated_author = "Updated Author"

      # Add a comment first
      :ok = UmyaSpreadsheet.add_comment(spreadsheet, sheet_name, cell_address, comment_text, author)

      # Update the comment text only
      :ok = UmyaSpreadsheet.update_comment(spreadsheet, sheet_name, cell_address, updated_text)

      # Verify text was updated but author remains the same
      {:ok, retrieved_text, retrieved_author} = UmyaSpreadsheet.get_comment(spreadsheet, sheet_name, cell_address)
      assert retrieved_text == updated_text
      assert retrieved_author == author

      # Update both text and author
      :ok = UmyaSpreadsheet.update_comment(spreadsheet, sheet_name, cell_address, updated_text, updated_author)

      # Verify both text and author were updated
      {:ok, retrieved_text, retrieved_author} = UmyaSpreadsheet.get_comment(spreadsheet, sheet_name, cell_address)
      assert retrieved_text == updated_text
      assert retrieved_author == updated_author
    end

    test "remove_comment", %{spreadsheet: spreadsheet} do
      sheet_name = "Sheet1"
      cell_address = "C3"
      comment_text = "Comment to be removed"
      author = "Test Author"

      # Add a comment
      :ok = UmyaSpreadsheet.add_comment(spreadsheet, sheet_name, cell_address, comment_text, author)

      # Verify comment exists
      assert {:ok, _, _} = UmyaSpreadsheet.get_comment(spreadsheet, sheet_name, cell_address)

      # Remove the comment
      :ok = UmyaSpreadsheet.remove_comment(spreadsheet, sheet_name, cell_address)

      # Verify comment is removed
      assert {:error, _} = UmyaSpreadsheet.get_comment(spreadsheet, sheet_name, cell_address)
    end

    test "has_comments and get_comments_count", %{spreadsheet: spreadsheet} do
      sheet_name = "Sheet1"

      # Initially no comments
      assert UmyaSpreadsheet.has_comments(spreadsheet, sheet_name) == false
      assert UmyaSpreadsheet.get_comments_count(spreadsheet, sheet_name) == 0

      # Add some comments
      :ok = UmyaSpreadsheet.add_comment(spreadsheet, sheet_name, "D4", "Comment 1", "Author 1")
      :ok = UmyaSpreadsheet.add_comment(spreadsheet, sheet_name, "E5", "Comment 2", "Author 2")

      # Check again
      assert UmyaSpreadsheet.has_comments(spreadsheet, sheet_name) == true
      assert UmyaSpreadsheet.get_comments_count(spreadsheet, sheet_name) == 2

      # Remove a comment
      :ok = UmyaSpreadsheet.remove_comment(spreadsheet, sheet_name, "D4")

      # Check counts again
      assert UmyaSpreadsheet.has_comments(spreadsheet, sheet_name) == true
      assert UmyaSpreadsheet.get_comments_count(spreadsheet, sheet_name) == 1

      # Remove the last comment
      :ok = UmyaSpreadsheet.remove_comment(spreadsheet, sheet_name, "E5")

      # No more comments
      assert UmyaSpreadsheet.has_comments(spreadsheet, sheet_name) == false
      assert UmyaSpreadsheet.get_comments_count(spreadsheet, sheet_name) == 0
    end
  end
end
