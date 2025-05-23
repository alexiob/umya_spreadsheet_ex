defmodule UmyaSpreadsheet.CommentFunctions do
  @moduledoc """
  Functions for working with cell comments in spreadsheets.
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaNative

  @doc """
  Adds a comment to a cell.

  ## Parameters

  * `spreadsheet` - The spreadsheet reference
  * `sheet_name` - Name of the worksheet
  * `cell_address` - Cell address in A1 notation (e.g., "A1", "B2")
  * `text` - Text content of the comment
  * `author` - Name of the comment author

  ## Examples

      iex> UmyaSpreadsheet.add_comment(spreadsheet, "Sheet1", "A1", "This is a comment", "John Doe")
      :ok

  """
  @spec add_comment(Spreadsheet.t(), String.t(), String.t(), String.t(), String.t()) :: :ok | {:error, atom()}
  def add_comment(%Spreadsheet{reference: ref}, sheet_name, cell_address, text, author) do
    UmyaNative.add_comment(ref, sheet_name, cell_address, text, author)
  end

  @doc """
  Gets the comment text and author from a cell.

  ## Parameters

  * `spreadsheet` - The spreadsheet reference
  * `sheet_name` - Name of the worksheet
  * `cell_address` - Cell address in A1 notation (e.g., "A1", "B2")

  ## Examples

      iex> UmyaSpreadsheet.get_comment(spreadsheet, "Sheet1", "A1")
      {:ok, "This is a comment", "John Doe"}

  """
  @spec get_comment(Spreadsheet.t(), String.t(), String.t()) :: {:ok, String.t(), String.t()} | {:error, atom()}
  def get_comment(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    UmyaNative.get_comment(ref, sheet_name, cell_address)
  end

  @doc """
  Updates an existing comment in a cell.

  ## Parameters

  * `spreadsheet` - The spreadsheet reference
  * `sheet_name` - Name of the worksheet
  * `cell_address` - Cell address in A1 notation (e.g., "A1", "B2")
  * `text` - New text content for the comment
  * `author` - Optional new author name (if nil, keeps the existing author)

  ## Examples

      iex> UmyaSpreadsheet.update_comment(spreadsheet, "Sheet1", "A1", "Updated comment")
      :ok

      iex> UmyaSpreadsheet.update_comment(spreadsheet, "Sheet1", "A1", "Updated comment", "Jane Smith")
      :ok

  """
  @spec update_comment(Spreadsheet.t(), String.t(), String.t(), String.t(), String.t() | nil) :: :ok | {:error, atom()}
  def update_comment(%Spreadsheet{reference: ref}, sheet_name, cell_address, text, author \\ nil) do
    UmyaNative.update_comment(ref, sheet_name, cell_address, text, author)
  end

  @doc """
  Removes a comment from a cell.

  ## Parameters

  * `spreadsheet` - The spreadsheet reference
  * `sheet_name` - Name of the worksheet
  * `cell_address` - Cell address in A1 notation (e.g., "A1", "B2")

  ## Examples

      iex> UmyaSpreadsheet.remove_comment(spreadsheet, "Sheet1", "A1")
      :ok

  """
  @spec remove_comment(Spreadsheet.t(), String.t(), String.t()) :: :ok | {:error, atom()}
  def remove_comment(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    UmyaNative.remove_comment(ref, sheet_name, cell_address)
  end

  @doc """
  Checks if a sheet has any comments.

  ## Parameters

  * `spreadsheet` - The spreadsheet reference
  * `sheet_name` - Name of the worksheet

  ## Examples

      iex> UmyaSpreadsheet.has_comments(spreadsheet, "Sheet1")
      true

  """
  @spec has_comments(Spreadsheet.t(), String.t()) :: boolean() | {:error, atom()}
  def has_comments(%Spreadsheet{reference: ref}, sheet_name) do
    UmyaNative.has_comments(ref, sheet_name)
  end

  @doc """
  Gets the number of comments in a sheet.

  ## Parameters

  * `spreadsheet` - The spreadsheet reference
  * `sheet_name` - Name of the worksheet

  ## Examples

      iex> UmyaSpreadsheet.get_comments_count(spreadsheet, "Sheet1")
      3

  """
  @spec get_comments_count(Spreadsheet.t(), String.t()) :: integer() | {:error, atom()}
  def get_comments_count(%Spreadsheet{reference: ref}, sheet_name) do
    UmyaNative.get_comments_count(ref, sheet_name)
  end
end
