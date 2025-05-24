defmodule UmyaSpreadsheetTest.CommentTest do
  use ExUnit.Case, async: true

  alias UmyaSpreadsheet

  setup do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()
    %{spreadsheet: spreadsheet}
  end

  test "add_comment adds a comment to a cell", %{spreadsheet: spreadsheet} do
    # Add a comment to a cell
    assert :ok =
             UmyaSpreadsheet.add_comment(
               spreadsheet,
               "Sheet1",
               "A1",
               "This is a test comment",
               "Test User"
             )

    # Verify the comment exists
    assert {:ok, true} == UmyaSpreadsheet.has_comments(spreadsheet, "Sheet1")
    assert {:ok, 1} == UmyaSpreadsheet.get_comments_count(spreadsheet, "Sheet1")

    # Get the comment
    assert {:ok, "This is a test comment", "Test User"} =
             UmyaSpreadsheet.get_comment(spreadsheet, "Sheet1", "A1")
  end

  test "update_comment changes an existing comment", %{spreadsheet: spreadsheet} do
    # Add a comment to a cell
    assert :ok =
             UmyaSpreadsheet.add_comment(
               spreadsheet,
               "Sheet1",
               "B2",
               "Original comment",
               "Original Author"
             )

    # Update the comment text only
    assert :ok =
             UmyaSpreadsheet.update_comment(
               spreadsheet,
               "Sheet1",
               "B2",
               "Updated comment"
             )

    # Verify the comment was updated (author remains the same)
    assert {:ok, "Updated comment", "Original Author"} =
             UmyaSpreadsheet.get_comment(spreadsheet, "Sheet1", "B2")

    # Update both text and author
    assert :ok =
             UmyaSpreadsheet.update_comment(
               spreadsheet,
               "Sheet1",
               "B2",
               "Fully updated comment",
               "New Author"
             )

    # Verify both text and author were updated
    assert {:ok, "Fully updated comment", "New Author"} =
             UmyaSpreadsheet.get_comment(spreadsheet, "Sheet1", "B2")
  end

  test "remove_comment removes a comment from a cell", %{spreadsheet: spreadsheet} do
    # Add a comment to a cell
    assert :ok =
             UmyaSpreadsheet.add_comment(
               spreadsheet,
               "Sheet1",
               "C3",
               "Comment to be removed",
               "Test User"
             )

    # Verify the comment exists
    assert {:ok, 1} == UmyaSpreadsheet.get_comments_count(spreadsheet, "Sheet1")

    # Remove the comment
    assert :ok =
             UmyaSpreadsheet.remove_comment(
               spreadsheet,
               "Sheet1",
               "C3"
             )

    # Verify the comment was removed
    assert {:error, _} = UmyaSpreadsheet.get_comment(spreadsheet, "Sheet1", "C3")
    assert {:ok, 0} == UmyaSpreadsheet.get_comments_count(spreadsheet, "Sheet1")
  end

  test "has_comments and get_comments_count work with multiple comments", %{spreadsheet: spreadsheet} do
    # Add multiple comments to different cells
    assert :ok =
             UmyaSpreadsheet.add_comment(
               spreadsheet,
               "Sheet1",
               "D1",
               "First comment",
               "User 1"
             )

    assert :ok =
             UmyaSpreadsheet.add_comment(
               spreadsheet,
               "Sheet1",
               "D2",
               "Second comment",
               "User 2"
             )

    assert :ok =
             UmyaSpreadsheet.add_comment(
               spreadsheet,
               "Sheet1",
               "D3",
               "Third comment",
               "User 3"
             )

    # Verify all comments exist
    assert {:ok, true} == UmyaSpreadsheet.has_comments(spreadsheet, "Sheet1")
    assert {:ok, 3} == UmyaSpreadsheet.get_comments_count(spreadsheet, "Sheet1")
  end

  test "get_comment returns error for non-existent comment", %{spreadsheet: spreadsheet} do
    # Try to get a comment that doesn't exist
    assert {:error, _} = UmyaSpreadsheet.get_comment(spreadsheet, "Sheet1", "Z10")
  end

  test "update_comment returns error for non-existent comment", %{spreadsheet: spreadsheet} do
    # Try to update a comment that doesn't exist
    assert {:error, _} =
             UmyaSpreadsheet.update_comment(
               spreadsheet,
               "Sheet1",
               "Z10",
               "This won't work"
             )
  end

  test "remove_comment returns error for non-existent comment", %{spreadsheet: spreadsheet} do
    # Try to remove a comment that doesn't exist
    assert {:error, _} = UmyaSpreadsheet.remove_comment(spreadsheet, "Sheet1", "Z10")
  end
end
