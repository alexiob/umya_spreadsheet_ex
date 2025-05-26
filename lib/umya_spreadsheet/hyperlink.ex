defmodule UmyaSpreadsheet.Hyperlink do
  @moduledoc """
  Functions for managing hyperlinks in Excel spreadsheets.

  This module provides a comprehensive API for working with hyperlinks in Excel cells,
  including support for:
  - Web URLs (http/https)
  - File paths (local files)
  - Internal worksheet references
  - Email addresses (mailto links)
  - Custom tooltip text
  - Complete hyperlink lifecycle management (add, get, update, remove)

  ## Examples

      # Add a web URL hyperlink
      {:ok, spreadsheet} = UmyaSpreadsheet.Hyperlink.add_hyperlink(
        spreadsheet,
        "Sheet1",
        "A1",
        "https://example.com",
        "Visit Example.com"
      )

      # Add an internal worksheet reference
      {:ok, spreadsheet} = UmyaSpreadsheet.Hyperlink.add_hyperlink(
        spreadsheet,
        "Sheet1",
        "B1",
        "Sheet2!A1",
        "Go to Sheet2",
        true
      )

      # Get hyperlink information
      {:ok, hyperlink_info} = UmyaSpreadsheet.Hyperlink.get_hyperlink(spreadsheet, "Sheet1", "A1")

      # Check if cell has hyperlink
      true = UmyaSpreadsheet.Hyperlink.has_hyperlink?(spreadsheet, "Sheet1", "A1")

      # Get all hyperlinks in worksheet
      {:ok, hyperlinks} = UmyaSpreadsheet.Hyperlink.get_all_hyperlinks(spreadsheet, "Sheet1")

      # Update existing hyperlink
      {:ok, spreadsheet} = UmyaSpreadsheet.Hyperlink.update_hyperlink(
        spreadsheet,
        "Sheet1",
        "A1",
        "https://newexample.com",
        "Updated tooltip"
      )

      # Remove hyperlink
      {:ok, spreadsheet} = UmyaSpreadsheet.Hyperlink.remove_hyperlink(spreadsheet, "Sheet1", "A1")
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaSpreadsheet.ErrorHandling
  alias UmyaNative

  @type hyperlink_info :: %{
          url: String.t(),
          tooltip: String.t() | nil,
          is_internal: boolean(),
          cell: String.t()
        }

  @doc """
  Adds a hyperlink to a cell.

  ## Parameters

  - `spreadsheet`: The spreadsheet reference
  - `sheet_name`: Name of the worksheet
  - `cell_address`: Cell address (e.g., "A1", "B2")
  - `url`: The hyperlink URL or reference
  - `tooltip`: Optional tooltip text (default: nil)
  - `is_internal`: Whether this is an internal worksheet reference (default: false)

  ## Returns

  - `{:ok, spreadsheet}` on success
  - `{:error, reason}` on failure

  ## Examples

      # Web URL
      {:ok, spreadsheet} = UmyaSpreadsheet.Hyperlink.add_hyperlink(
        spreadsheet,
        "Sheet1",
        "A1",
        "https://example.com",
        "Visit our website"
      )

      # Email link
      {:ok, spreadsheet} = UmyaSpreadsheet.Hyperlink.add_hyperlink(
        spreadsheet,
        "Sheet1",
        "B1",
        "mailto:contact@example.com",
        "Send email"
      )

      # File path
      {:ok, spreadsheet} = UmyaSpreadsheet.Hyperlink.add_hyperlink(
        spreadsheet,
        "Sheet1",
        "C1",
        "file:///path/to/document.pdf",
        "Open document"
      )

      # Internal reference
      {:ok, spreadsheet} = UmyaSpreadsheet.Hyperlink.add_hyperlink(
        spreadsheet,
        "Sheet1",
        "D1",
        "Sheet2!A1",
        "Go to Sheet2",
        true
      )
  """
  @spec add_hyperlink(
          Spreadsheet.t(),
          String.t(),
          String.t(),
          String.t(),
          String.t() | nil,
          boolean()
        ) :: :ok | {:error, atom()}
  def add_hyperlink(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_address,
        url,
        tooltip \\ nil,
        is_internal \\ false
      ) do
    UmyaNative.add_hyperlink(ref, sheet_name, cell_address, url, tooltip, is_internal)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets hyperlink information from a cell.

  ## Parameters

  - `spreadsheet`: The spreadsheet reference
  - `sheet_name`: Name of the worksheet
  - `cell_address`: Cell address (e.g., "A1", "B2")

  ## Returns

  - `{:ok, hyperlink_info}` if hyperlink exists
  - `{:error, :not_found}` if no hyperlink exists
  - `{:error, reason}` on other failures

  ## Examples

      {:ok, info} = UmyaSpreadsheet.Hyperlink.get_hyperlink(spreadsheet, "Sheet1", "A1")
      # Returns: %{url: "https://example.com", tooltip: "Visit our website", is_internal: false, cell: "A1"}
  """
  @spec get_hyperlink(Spreadsheet.t(), String.t(), String.t()) ::
          {:ok, hyperlink_info()} | {:error, atom()}
  def get_hyperlink(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    UmyaNative.get_hyperlink(ref, sheet_name, cell_address)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Removes a hyperlink from a cell.

  The cell value and formatting are preserved, only the hyperlink is removed.

  ## Parameters

  - `spreadsheet`: The spreadsheet reference
  - `sheet_name`: Name of the worksheet
  - `cell_address`: Cell address (e.g., "A1", "B2")

  ## Returns

  - `{:ok, spreadsheet}` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.Hyperlink.remove_hyperlink(spreadsheet, "Sheet1", "A1")
  """
  @spec remove_hyperlink(Spreadsheet.t(), String.t(), String.t()) ::
          :ok | {:error, atom()}
  def remove_hyperlink(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    UmyaNative.remove_hyperlink(ref, sheet_name, cell_address)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Checks if a specific cell has a hyperlink.

  ## Parameters

  - `spreadsheet`: The spreadsheet reference
  - `sheet_name`: Name of the worksheet
  - `cell_address`: Cell address (e.g., "A1", "B2")

  ## Returns

  - `true` if the cell has a hyperlink
  - `false` if the cell doesn't have a hyperlink
  - `{:error, reason}` on failure

  ## Examples

      true = UmyaSpreadsheet.Hyperlink.has_hyperlink?(spreadsheet, "Sheet1", "A1")
      false = UmyaSpreadsheet.Hyperlink.has_hyperlink?(spreadsheet, "Sheet1", "B1")
  """
  @spec has_hyperlink?(Spreadsheet.t(), String.t(), String.t()) ::
          {:ok, boolean()} | {:error, atom()}
  def has_hyperlink?(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    UmyaNative.has_hyperlink(ref, sheet_name, cell_address)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Checks if a worksheet contains any hyperlinks.

  ## Parameters

  - `spreadsheet`: The spreadsheet reference
  - `sheet_name`: Name of the worksheet

  ## Returns

  - `true` if the worksheet has any hyperlinks
  - `false` if the worksheet has no hyperlinks
  - `{:error, reason}` on failure

  ## Examples

      true = UmyaSpreadsheet.Hyperlink.has_hyperlinks?(spreadsheet, "Sheet1")
      false = UmyaSpreadsheet.Hyperlink.has_hyperlinks?(spreadsheet, "EmptySheet")
  """
  @spec has_hyperlinks?(Spreadsheet.t(), String.t()) :: {:ok, boolean()} | {:error, atom()}
  def has_hyperlinks?(%Spreadsheet{reference: ref}, sheet_name) do
    UmyaNative.has_hyperlinks(ref, sheet_name)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets all hyperlinks from a worksheet.

  ## Parameters

  - `spreadsheet`: The spreadsheet reference
  - `sheet_name`: Name of the worksheet

  ## Returns

  - `{:ok, hyperlinks}` where hyperlinks is a list of hyperlink_info maps
  - `{:error, reason}` on failure

  ## Examples

      {:ok, hyperlinks} = UmyaSpreadsheet.Hyperlink.get_all_hyperlinks(spreadsheet, "Sheet1")
      # Returns: [
      #   %{url: "https://example.com", tooltip: "Website", is_internal: false, cell: "A1"},
      #   %{url: "Sheet2!A1", tooltip: "Internal link", is_internal: true, cell: "B1"}
      # ]
  """
  @spec get_all_hyperlinks(Spreadsheet.t(), String.t()) ::
          {:ok, [hyperlink_info()]} | {:error, atom()}
  def get_all_hyperlinks(%Spreadsheet{reference: ref}, sheet_name) do
    UmyaNative.get_hyperlinks(ref, sheet_name)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Updates an existing hyperlink in a cell.

  If no hyperlink exists in the cell, this function will add a new one.

  ## Parameters

  - `spreadsheet`: The spreadsheet reference
  - `sheet_name`: Name of the worksheet
  - `cell_address`: Cell address (e.g., "A1", "B2")
  - `url`: The new hyperlink URL or reference
  - `tooltip`: Optional new tooltip text (default: nil)
  - `is_internal`: Whether this is an internal worksheet reference (default: false)

  ## Returns

  - `{:ok, spreadsheet}` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.Hyperlink.update_hyperlink(
        spreadsheet,
        "Sheet1",
        "A1",
        "https://newexample.com",
        "Updated website link"
      )
  """
  @spec update_hyperlink(
          Spreadsheet.t(),
          String.t(),
          String.t(),
          String.t(),
          String.t() | nil,
          boolean()
        ) :: :ok | {:error, atom()}
  def update_hyperlink(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_address,
        url,
        tooltip \\ nil,
        is_internal \\ false
      ) do
    UmyaNative.update_hyperlink(ref, sheet_name, cell_address, url, tooltip, is_internal)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Bulk adds multiple hyperlinks to a worksheet.

  ## Parameters

  - `spreadsheet`: The spreadsheet reference
  - `sheet_name`: Name of the worksheet
  - `hyperlinks`: List of hyperlink specifications as tuples or maps

  ## Hyperlink specifications

  Each hyperlink can be specified as:
  - `{cell_address, url}` - Basic hyperlink without tooltip
  - `{cell_address, url, tooltip}` - Hyperlink with tooltip
  - `{cell_address, url, tooltip, is_internal}` - Full specification
  - `%{cell: cell_address, url: url, tooltip: tooltip, is_internal: is_internal}` - Map format

  ## Returns

  - `{:ok, spreadsheet}` on success
  - `{:error, reason}` on failure

  ## Examples

      hyperlinks = [
        {"A1", "https://example.com", "Website"},
        {"B1", "mailto:test@example.com", "Email"},
        {"C1", "Sheet2!A1", "Internal link", true}
      ]

      {:ok, spreadsheet} = UmyaSpreadsheet.Hyperlink.add_bulk_hyperlinks(
        spreadsheet,
        "Sheet1",
        hyperlinks
      )
  """
  @spec add_bulk_hyperlinks(
          Spreadsheet.t(),
          String.t(),
          [
            {String.t(), String.t()}
            | {String.t(), String.t(), String.t() | nil}
            | {String.t(), String.t(), String.t() | nil, boolean()}
            | hyperlink_info()
          ]
        ) :: :ok | {:error, atom()}
  def add_bulk_hyperlinks(spreadsheet, sheet_name, hyperlinks) do
    result =
      Enum.reduce_while(hyperlinks, :ok, fn hyperlink_spec, :ok ->
        case process_hyperlink_spec(spreadsheet, sheet_name, hyperlink_spec) do
          :ok -> {:cont, :ok}
          {:error, reason} -> {:halt, {:error, reason}}
        end
      end)

    result
  end

  # Private helper function to process different hyperlink specification formats
  defp process_hyperlink_spec(spreadsheet, sheet_name, {cell, url}) do
    add_hyperlink(spreadsheet, sheet_name, cell, url)
  end

  defp process_hyperlink_spec(spreadsheet, sheet_name, {cell, url, tooltip}) do
    add_hyperlink(spreadsheet, sheet_name, cell, url, tooltip)
  end

  defp process_hyperlink_spec(spreadsheet, sheet_name, {cell, url, tooltip, is_internal}) do
    add_hyperlink(spreadsheet, sheet_name, cell, url, tooltip, is_internal)
  end

  defp process_hyperlink_spec(spreadsheet, sheet_name, %{cell: cell, url: url} = spec) do
    tooltip = Map.get(spec, :tooltip)
    is_internal = Map.get(spec, :is_internal, false)
    add_hyperlink(spreadsheet, sheet_name, cell, url, tooltip, is_internal)
  end

  defp process_hyperlink_spec(_spreadsheet, _sheet_name, invalid_spec) do
    {:error, "Invalid hyperlink specification: #{inspect(invalid_spec)}"}
  end

  @doc """
  Removes all hyperlinks from a worksheet.

  Cell values and formatting are preserved, only hyperlinks are removed.

  ## Parameters

  - `spreadsheet`: The spreadsheet reference
  - `sheet_name`: Name of the worksheet

  ## Returns

  - `{:ok, spreadsheet}` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.Hyperlink.remove_all_hyperlinks(spreadsheet, "Sheet1")
  """
  @spec remove_all_hyperlinks(Spreadsheet.t(), String.t()) ::
          :ok | {:error, atom()}
  def remove_all_hyperlinks(spreadsheet, sheet_name) do
    case get_all_hyperlinks(spreadsheet, sheet_name) do
      {:ok, hyperlinks} ->
        result =
          Enum.reduce_while(hyperlinks, :ok, fn %{cell: cell}, :ok ->
            case remove_hyperlink(spreadsheet, sheet_name, cell) do
              :ok -> {:cont, :ok}
              {:error, reason} -> {:halt, {:error, reason}}
            end
          end)

        result

      {:error, reason} ->
        {:error, reason}
    end
  end

  @doc """
  Counts the number of hyperlinks in a worksheet.

  ## Parameters

  - `spreadsheet`: The spreadsheet reference
  - `sheet_name`: Name of the worksheet

  ## Returns

  - `{:ok, count}` where count is the number of hyperlinks
  - `{:error, reason}` on failure

  ## Examples

      {:ok, 5} = UmyaSpreadsheet.Hyperlink.count_hyperlinks(spreadsheet, "Sheet1")
  """
  @spec count_hyperlinks(Spreadsheet.t(), String.t()) ::
          {:ok, non_neg_integer()} | {:error, atom()}
  def count_hyperlinks(spreadsheet, sheet_name) do
    case get_all_hyperlinks(spreadsheet, sheet_name) do
      {:ok, hyperlinks} -> {:ok, length(hyperlinks)}
      {:error, reason} -> {:error, reason}
    end
  end
end
