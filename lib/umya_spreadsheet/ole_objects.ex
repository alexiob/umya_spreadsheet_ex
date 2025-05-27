defmodule UmyaSpreadsheet.OleObjects do
  @moduledoc """
  Functions for working with OLE (Object Linking and Embedding) objects in a spreadsheet.

  OLE objects allow you to embed documents and files from other applications into your
  Excel spreadsheets, such as Word documents, PowerPoint presentations, PDF files, and more.

  ## Overview

  This module provides functionality to:
  - Create and manage collections of OLE objects
  - Add embedded objects to worksheets
  - Load and save OLE object data
  - Configure object properties and settings
  - Check object formats and requirements

  ## Examples

      # Create a new OLE objects collection
      {:ok, ole_objects} = UmyaSpreadsheet.OleObjects.new_ole_objects()

      # Add the collection to a worksheet
      :ok = UmyaSpreadsheet.OleObjects.set_ole_objects(spreadsheet, "Sheet1", ole_objects)

      # Create an OLE object from a file
      {:ok, ole_object} = UmyaSpreadsheet.OleObjects.new_ole_object_from_file("document.docx")

      # Add the object to the collection
      :ok = UmyaSpreadsheet.OleObjects.add_ole_object(ole_objects, ole_object)
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaSpreadsheet.ErrorHandling
  alias UmyaNative

  # OLE Objects Collection Management

  @doc """
  Creates a new OLE objects collection.

  ## Returns

  - `{:ok, ole_objects}` - A new OLE objects collection resource
  - `{:error, reason}` on failure

  ## Examples

      {:ok, ole_objects} = UmyaSpreadsheet.OleObjects.new_ole_objects()
  """
  def new_ole_objects do
    UmyaNative.create_ole_objects()
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the OLE objects collection from a worksheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the worksheet

  ## Returns

  - `{:ok, ole_objects}` - The OLE objects collection from the worksheet
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, ole_objects} = UmyaSpreadsheet.OleObjects.get_ole_objects_from_worksheet(spreadsheet, "Sheet1")
  """
  def get_ole_objects_from_worksheet(%Spreadsheet{reference: ref}, sheet_name) do
    UmyaNative.get_ole_objects(ref, sheet_name)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the OLE objects collection for a worksheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the worksheet
  - `ole_objects` - The OLE objects collection resource

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, ole_objects} = UmyaSpreadsheet.OleObjects.new_ole_objects()
      :ok = UmyaSpreadsheet.OleObjects.set_ole_objects_to_worksheet(spreadsheet, "Sheet1", ole_objects)
  """
  def set_ole_objects_to_worksheet(%Spreadsheet{reference: ref}, sheet_name, ole_objects) do
    UmyaNative.set_ole_objects(ref, sheet_name, ole_objects)
    |> ErrorHandling.standardize_result()
  end

  # Individual OLE Object Management

  @doc """
  Creates a new OLE object.

  ## Returns

  - `{:ok, ole_object}` - A new OLE object resource
  - `{:error, reason}` on failure

  ## Examples

      {:ok, ole_object} = UmyaSpreadsheet.OleObjects.new_ole_object()
  """
  def new_ole_object do
    UmyaNative.create_ole_object()
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Creates a new OLE object and loads data from a file.

  ## Parameters

  - `file_path` - Path to the file to embed

  ## Returns

  - `{:ok, ole_object}` - A new OLE object with file data loaded
  - `{:error, reason}` on failure

  ## Examples

      {:ok, ole_object} = UmyaSpreadsheet.OleObjects.new_ole_object_from_file("document.docx")
  """
  def new_ole_object_from_file(file_path) do
    # Based on umya_native.ex, this should be load_ole_object_from_file but it takes different params
    # We'll need to create a simple wrapper that determines the prog_id from file extension
    extension = file_path |> Path.extname() |> String.trim_leading(".")

    prog_id =
      case extension do
        "docx" -> "Word.Document.12"
        "xlsx" -> "Excel.Sheet.12"
        "pptx" -> "PowerPoint.Show.12"
        _ -> "Package"
      end

    UmyaNative.load_ole_object_from_file(file_path, prog_id)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Creates a new OLE object with binary data.

  ## Parameters

  - `data` - Binary data to embed
  - `file_extension` - File extension (e.g., "docx", "pptx")

  ## Returns

  - `{:ok, ole_object}` - A new OLE object with the provided data
  - `{:error, reason}` on failure

  ## Examples

      data = File.read!("document.docx")
      {:ok, ole_object} = UmyaSpreadsheet.OleObjects.new_ole_object_with_data(data, "docx")
  """
  def new_ole_object_with_data(data, file_extension) do
    prog_id =
      case file_extension do
        "docx" -> "Word.Document.12"
        "xlsx" -> "Excel.Sheet.12"
        "pptx" -> "PowerPoint.Show.12"
        _ -> "Package"
      end

    try do
      UmyaNative.create_ole_object_with_data(prog_id, file_extension, data)
      |> ErrorHandling.standardize_result()
    rescue
      ArgumentError -> {:error, "Function not implemented"}
    end
  end

  # Collection Operations

  @doc """
  Adds an OLE object to a collection.

  ## Parameters

  - `ole_objects` - The OLE objects collection
  - `ole_object` - The OLE object to add

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, ole_objects} = UmyaSpreadsheet.OleObjects.new_ole_objects()
      {:ok, ole_object} = UmyaSpreadsheet.OleObjects.new_ole_object_from_file("document.docx")
      :ok = UmyaSpreadsheet.OleObjects.add_ole_object(ole_objects, ole_object)
  """
  def add_ole_object(ole_objects, ole_object) do
    UmyaNative.add_ole_object(ole_objects, ole_object)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Lists all OLE objects in a collection.

  ## Parameters

  - `ole_objects` - The OLE objects collection

  ## Returns

  - `{:ok, objects_list}` - List of OLE object resources
  - `{:error, reason}` on failure

  ## Examples

      {:ok, ole_objects} = UmyaSpreadsheet.OleObjects.get_ole_objects_from_worksheet(spreadsheet, "Sheet1")
      {:ok, objects_list} = UmyaSpreadsheet.OleObjects.list_ole_objects(ole_objects)
  """
  def list_ole_objects(ole_objects) do
    UmyaNative.get_ole_object_list(ole_objects)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the count of OLE objects in a collection.

  ## Parameters

  - `ole_objects` - The OLE objects collection

  ## Returns

  - `{:ok, count}` - Number of OLE objects in the collection
  - `{:error, reason}` on failure

  ## Examples

      {:ok, ole_objects} = UmyaSpreadsheet.OleObjects.get_ole_objects_from_worksheet(spreadsheet, "Sheet1")
      {:ok, count} = UmyaSpreadsheet.OleObjects.get_ole_objects_count(ole_objects)
  """
  def get_ole_objects_count(ole_objects) do
    UmyaNative.count_ole_objects(ole_objects)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Checks if an OLE objects collection has any objects.

  ## Parameters

  - `ole_objects` - The OLE objects collection

  ## Returns

  - `{:ok, boolean}` - true if collection has objects, false otherwise
  - `{:error, reason}` on failure

  ## Examples

      {:ok, ole_objects} = UmyaSpreadsheet.OleObjects.get_ole_objects_from_worksheet(spreadsheet, "Sheet1")
      {:ok, has_objects} = UmyaSpreadsheet.OleObjects.has_ole_objects(ole_objects)
  """
  def has_ole_objects(ole_objects) do
    UmyaNative.has_ole_objects(ole_objects)
    |> ErrorHandling.standardize_result()
  end

  # OLE Object Properties

  @doc """
  Gets the properties of an OLE object.

  ## Parameters

  - `ole_object` - The OLE object resource

  ## Returns

  - `{:ok, properties}` - The embedded object properties resource
  - `{:error, reason}` on failure

  ## Examples

      {:ok, ole_object} = UmyaSpreadsheet.OleObjects.new_ole_object_from_file("document.docx")
      {:ok, properties} = UmyaSpreadsheet.OleObjects.get_ole_object_properties(ole_object)
  """
  def get_ole_object_properties(ole_object) do
    UmyaNative.get_ole_object_properties(ole_object)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the properties for an OLE object.

  ## Parameters

  - `ole_object` - The OLE object resource
  - `properties` - The embedded object properties resource

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, ole_object} = UmyaSpreadsheet.OleObjects.new_ole_object()
      {:ok, properties} = UmyaSpreadsheet.OleObjects.new_embedded_object_properties()
      :ok = UmyaSpreadsheet.OleObjects.set_ole_object_properties(ole_object, properties)
  """
  def set_ole_object_properties(ole_object, properties) do
    UmyaNative.set_ole_object_properties(ole_object, properties)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Creates new embedded object properties.

  ## Returns

  - `{:ok, properties}` - A new embedded object properties resource
  - `{:error, reason}` on failure

  ## Examples

      {:ok, properties} = UmyaSpreadsheet.OleObjects.new_embedded_object_properties()
  """
  def new_embedded_object_properties do
    UmyaNative.create_embedded_object_properties()
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the 'requires' attribute from an OLE object.

  ## Parameters

  - `ole_object` - The OLE object resource

  ## Returns

  - `{:ok, requires}` - The requires string value
  - `{:error, reason}` on failure

  ## Examples

      {:ok, ole_object} = UmyaSpreadsheet.OleObjects.new_ole_object_from_file("document.docx")
      {:ok, requires} = UmyaSpreadsheet.OleObjects.get_ole_object_requires(ole_object)
  """
  def get_ole_object_requires(ole_object) do
    UmyaNative.get_ole_object_requires(ole_object)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the 'requires' attribute for an OLE object.

  ## Parameters

  - `ole_object` - The OLE object resource
  - `requires` - The requires string value

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, ole_object} = UmyaSpreadsheet.OleObjects.new_ole_object()
      :ok = UmyaSpreadsheet.OleObjects.set_ole_object_requires(ole_object, "application")
  """
  def set_ole_object_requires(ole_object, requires) do
    UmyaNative.set_ole_object_requires(ole_object, requires)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the ProgID from an OLE object.

  ## Parameters

  - `ole_object` - The OLE object resource

  ## Returns

  - `{:ok, prog_id}` - The ProgID string value
  - `{:error, reason}` on failure

  ## Examples

      {:ok, ole_object} = UmyaSpreadsheet.OleObjects.new_ole_object_from_file("document.docx")
      {:ok, prog_id} = UmyaSpreadsheet.OleObjects.get_ole_object_prog_id(ole_object)
  """
  def get_ole_object_prog_id(ole_object) do
    UmyaNative.get_ole_object_prog_id(ole_object)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the ProgID for an OLE object.

  ## Parameters

  - `ole_object` - The OLE object resource
  - `prog_id` - The ProgID string value

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, ole_object} = UmyaSpreadsheet.OleObjects.new_ole_object()
      :ok = UmyaSpreadsheet.OleObjects.set_ole_object_prog_id(ole_object, "Word.Document.12")
  """
  def set_ole_object_prog_id(ole_object, prog_id) do
    UmyaNative.set_ole_object_prog_id(ole_object, prog_id)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the file extension from an OLE object.

  ## Parameters

  - `ole_object` - The OLE object resource

  ## Returns

  - `{:ok, extension}` - The file extension string
  - `{:error, reason}` on failure

  ## Examples

      {:ok, ole_object} = UmyaSpreadsheet.OleObjects.new_ole_object_from_file("document.docx")
      {:ok, extension} = UmyaSpreadsheet.OleObjects.get_ole_object_extension(ole_object)
  """
  def get_ole_object_extension(ole_object) do
    UmyaNative.get_ole_object_extension(ole_object)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the file extension for an OLE object.

  ## Parameters

  - `ole_object` - The OLE object resource
  - `extension` - The file extension string

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, ole_object} = UmyaSpreadsheet.OleObjects.new_ole_object()
      :ok = UmyaSpreadsheet.OleObjects.set_ole_object_extension(ole_object, "docx")
  """
  def set_ole_object_extension(ole_object, extension) do
    UmyaNative.set_ole_object_extension(ole_object, extension)
    |> ErrorHandling.standardize_result()
  end

  # Embedded Object Properties

  @doc """
  Gets the ProgID from embedded object properties.

  ## Parameters

  - `properties` - The embedded object properties resource

  ## Returns

  - `{:ok, prog_id}` - The ProgID string value
  - `{:error, reason}` on failure

  ## Examples

      {:ok, properties} = UmyaSpreadsheet.OleObjects.get_ole_object_properties(ole_object)
      {:ok, prog_id} = UmyaSpreadsheet.OleObjects.get_embedded_object_prog_id(properties)
  """
  def get_embedded_object_prog_id(properties) do
    UmyaNative.get_embedded_object_prog_id(properties)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the ProgID for embedded object properties.

  ## Parameters

  - `properties` - The embedded object properties resource
  - `prog_id` - The ProgID string value

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, properties} = UmyaSpreadsheet.OleObjects.new_embedded_object_properties()
      :ok = UmyaSpreadsheet.OleObjects.set_embedded_object_prog_id(properties, "Word.Document.12")
  """
  def set_embedded_object_prog_id(properties, prog_id) do
    UmyaNative.set_embedded_object_prog_id(properties, prog_id)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the shape ID from embedded object properties.

  ## Parameters

  - `properties` - The embedded object properties resource

  ## Returns

  - `{:ok, shape_id}` - The shape ID number
  - `{:error, reason}` on failure

  ## Examples

      {:ok, properties} = UmyaSpreadsheet.OleObjects.get_ole_object_properties(ole_object)
      {:ok, shape_id} = UmyaSpreadsheet.OleObjects.get_embedded_object_shape_id(properties)
  """
  def get_embedded_object_shape_id(properties) do
    UmyaNative.get_embedded_object_shape_id(properties)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the shape ID for embedded object properties.

  ## Parameters

  - `properties` - The embedded object properties resource
  - `shape_id` - The shape ID number

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, properties} = UmyaSpreadsheet.OleObjects.new_embedded_object_properties()
      :ok = UmyaSpreadsheet.OleObjects.set_embedded_object_shape_id(properties, 1)
  """
  def set_embedded_object_shape_id(properties, shape_id) do
    UmyaNative.set_embedded_object_shape_id(properties, shape_id)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the binary data from an OLE object.

  ## Parameters

  - `ole_object` - The OLE object resource

  ## Returns

  - `{:ok, data}` - The binary data
  - `{:error, reason}` on failure

  ## Examples

      {:ok, ole_object} = UmyaSpreadsheet.OleObjects.new_ole_object_from_file("document.docx")
      {:ok, data} = UmyaSpreadsheet.OleObjects.get_ole_object_data(ole_object)
  """
  def get_ole_object_data(ole_object) do
    UmyaNative.get_ole_object_data(ole_object)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the binary data for an OLE object.

  ## Parameters

  - `ole_object` - The OLE object resource
  - `data` - The binary data to set

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      data = File.read!("document.docx")
      {:ok, ole_object} = UmyaSpreadsheet.OleObjects.new_ole_object()
      :ok = UmyaSpreadsheet.OleObjects.set_ole_object_data(ole_object, data)
  """
  def set_ole_object_data(ole_object, data) do
    try do
      UmyaNative.set_ole_object_data(ole_object, data)
      |> ErrorHandling.standardize_result()
    rescue
      ArgumentError -> {:error, "Function not implemented"}
    end
  end

  # File Operations

  @doc """
  Loads an OLE object from a file.
  Note: This function signature differs from the underlying NIF which takes (file_path, prog_id).

  ## Parameters

  - `ole_object` - This parameter is ignored for this function
  - `file_path` - Path to the file to load

  ## Returns

  - `{:ok, ole_object}` - A new OLE object with file data loaded
  - `{:error, reason}` on failure

  ## Examples

      # This creates a new OLE object rather than loading into an existing one
      {:ok, ole_object} = UmyaSpreadsheet.OleObjects.load_ole_object_from_file(nil, "document.docx")
  """
  def load_ole_object_from_file(_ole_object, file_path) do
    extension = file_path |> Path.extname() |> String.trim_leading(".")

    prog_id =
      case extension do
        "docx" -> "Word.Document.12"
        "xlsx" -> "Excel.Sheet.12"
        "pptx" -> "PowerPoint.Show.12"
        _ -> "Package"
      end

    UmyaNative.load_ole_object_from_file(file_path, prog_id)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Saves an OLE object to a file.

  ## Parameters

  - `ole_object` - The OLE object resource
  - `file_path` - Path where the file should be saved

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, ole_object} = UmyaSpreadsheet.OleObjects.new_ole_object_from_file("input.docx")
      :ok = UmyaSpreadsheet.OleObjects.save_ole_object_to_file(ole_object, "output.docx")
  """
  def save_ole_object_to_file(ole_object, file_path) do
    UmyaNative.save_ole_object_to_file(ole_object, file_path)
    |> ErrorHandling.standardize_result()
  end

  # Format Checking

  @doc """
  Checks if an OLE object is in binary format.

  ## Parameters

  - `ole_object` - The OLE object resource

  ## Returns

  - `{:ok, boolean}` - true if in binary format, false otherwise
  - `{:error, reason}` on failure

  ## Examples

      {:ok, ole_object} = UmyaSpreadsheet.OleObjects.new_ole_object_from_file("document.docx")
      {:ok, is_binary} = UmyaSpreadsheet.OleObjects.is_ole_object_binary_format(ole_object)
  """
  def is_ole_object_binary_format(ole_object) do
    UmyaNative.is_ole_object_binary(ole_object)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Checks if an OLE object is in Excel format.

  ## Parameters

  - `ole_object` - The OLE object resource

  ## Returns

  - `{:ok, boolean}` - true if in Excel format, false otherwise
  - `{:error, reason}` on failure

  ## Examples

      {:ok, ole_object} = UmyaSpreadsheet.OleObjects.new_ole_object_from_file("spreadsheet.xlsx")
      {:ok, is_excel} = UmyaSpreadsheet.OleObjects.is_ole_object_excel_format(ole_object)
  """
  def is_ole_object_excel_format(ole_object) do
    UmyaNative.is_ole_object_excel(ole_object)
    |> ErrorHandling.standardize_result()
  end

  # Helper Functions

  @doc """
  Determines the ProgID from a file extension.

  ## Parameters

  - `extension` - The file extension (without the dot)

  ## Returns

  - `string` - The appropriate ProgID for the extension

  ## Examples

      "Word.Document.12" = UmyaSpreadsheet.OleObjects.determine_prog_id("docx")
      "Excel.Sheet.12" = UmyaSpreadsheet.OleObjects.determine_prog_id("xlsx")
  """
  def determine_prog_id(extension) do
    case extension do
      "docx" -> "Word.Document.12"
      "xlsx" -> "Excel.Sheet.12"
      "pptx" -> "PowerPoint.Show.12"
      "pdf" -> "AcroExch.Document"
      "txt" -> "txtfile"
      _ -> "Package"
    end
  end
end
