defmodule UmyaSpreadsheet.ImageFunctions do
  @moduledoc """
  Functions for working with images in a spreadsheet.
  """
  
  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaNative

  @doc """
  Adds an image to a spreadsheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1") where the top-left corner of the image should be placed
  - `image_path` - The path to the image file

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.ImageFunctions.add_image(spreadsheet, "Sheet1", "C5", "path/to/image.png")
  """
  def add_image(%Spreadsheet{reference: ref}, sheet_name, cell_address, image_path) do
    case UmyaNative.add_image(ref, sheet_name, cell_address, image_path) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Extracts an image from a spreadsheet and saves it to a file.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1") where the image is located
  - `output_path` - The path where the image should be saved

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.ImageFunctions.download_image(spreadsheet, "Sheet1", "C5", "extracted_image.png")
  """
  def download_image(%Spreadsheet{reference: ref}, sheet_name, cell_address, output_path) do
    case UmyaNative.download_image(ref, sheet_name, cell_address, output_path) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end

  @doc """
  Changes an existing image in a spreadsheet to a new image.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1") where the image is located
  - `new_image_path` - The path to the new image file

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.ImageFunctions.change_image(spreadsheet, "Sheet1", "C5", "path/to/new_image.png")
  """
  def change_image(%Spreadsheet{reference: ref}, sheet_name, cell_address, new_image_path) do
    case UmyaNative.change_image(ref, sheet_name, cell_address, new_image_path) do
      {:ok, :ok} -> :ok
      :ok -> :ok
      result -> result
    end
  end
end
