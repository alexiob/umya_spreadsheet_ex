defmodule UmyaSpreadsheet.ImageFunctions do
  @moduledoc """
  Functions for working with images in a spreadsheet.
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaSpreadsheet.ErrorHandling
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
    UmyaNative.add_image(ref, sheet_name, cell_address, image_path)
    |> ErrorHandling.standardize_result()
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
    UmyaNative.download_image(ref, sheet_name, cell_address, output_path)
    |> ErrorHandling.standardize_result()
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
    UmyaNative.change_image(ref, sheet_name, cell_address, new_image_path)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the dimensions (width, height) of an image in a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1") where the image is located

  ## Returns

  - `{:ok, {width, height}}` where width and height are in pixels
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, {width, height}} = UmyaSpreadsheet.ImageFunctions.get_image_dimensions(spreadsheet, "Sheet1", "C5")
  """
  def get_image_dimensions(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    UmyaNative.get_image_dimensions(ref, sheet_name, cell_address)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Lists all images in a sheet with their positions and names.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet

  ## Returns

  - `{:ok, [{coordinate, image_name}, ...]}` with a list of image coordinates and names
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, images} = UmyaSpreadsheet.ImageFunctions.list_images(spreadsheet, "Sheet1")

      # Print all images in the sheet
      Enum.each(images, fn {coord, img_name} ->
        IO.puts("Image '\#{img_name}' at position \#{coord}")
      end)
  """
  def list_images(%Spreadsheet{reference: ref}, sheet_name) do
    UmyaNative.list_images(ref, sheet_name)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets comprehensive information about an image in a spreadsheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1") where the image is located

  ## Returns

  - `{:ok, {name, position, width, height}}` with the image details
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, {name, pos, width, height}} = UmyaSpreadsheet.ImageFunctions.get_image_info(spreadsheet, "Sheet1", "C5")

      # Print the image information
      IO.puts("Image '" <> name <>"' at position " <> pos <>", dimensions: "<>width{}"x"<>height<>" pixels")
  """
  def get_image_info(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    UmyaNative.get_image_info(ref, sheet_name, cell_address)
    |> ErrorHandling.standardize_result()
  end
end
