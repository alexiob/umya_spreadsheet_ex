defmodule UmyaSpreadsheet.OleObjectsTest do
  use ExUnit.Case, async: true
  doctest UmyaSpreadsheet.OleObjects

  alias UmyaSpreadsheet.OleObjects

  describe "OLE Objects collection management" do
    test "can create and manage OLE objects collection" do
      # Create a new OLE objects collection
      assert {:ok, ole_objects} = UmyaSpreadsheet.new_ole_objects()

      # Initially should be empty
      assert {:ok, 0} = UmyaSpreadsheet.get_ole_objects_count(ole_objects)
      assert {:ok, false} = UmyaSpreadsheet.has_ole_objects(ole_objects)
      assert {:ok, []} = UmyaSpreadsheet.list_ole_objects(ole_objects)
    end

    test "can add OLE objects to collection" do
      # Create collection and object
      assert {:ok, ole_objects} = UmyaSpreadsheet.new_ole_objects()
      assert {:ok, ole_object} = UmyaSpreadsheet.new_ole_object()

      # Add object to collection
      assert :ok = UmyaSpreadsheet.add_ole_object(ole_objects, ole_object)

      # Collection should now have one object
      assert {:ok, 1} = UmyaSpreadsheet.get_ole_objects_count(ole_objects)
      assert {:ok, true} = UmyaSpreadsheet.has_ole_objects(ole_objects)
      assert {:ok, objects_list} = UmyaSpreadsheet.list_ole_objects(ole_objects)
      assert length(objects_list) == 1
    end
  end

  describe "OLE Object properties" do
    test "can create and set embedded object properties" do
      # Create OLE object and properties
      assert {:ok, ole_object} = UmyaSpreadsheet.new_ole_object()
      assert {:ok, properties} = UmyaSpreadsheet.new_embedded_object_properties()

      # Set properties for the OLE object
      assert :ok = UmyaSpreadsheet.set_ole_object_properties(ole_object, properties)

      # Get properties back
      assert {:ok, retrieved_properties} = UmyaSpreadsheet.get_ole_object_properties(ole_object)
      assert is_reference(retrieved_properties)
    end

    test "can set and get property values" do
      {:ok, ole_object} = UmyaSpreadsheet.new_ole_object()

      # Test requires property
      assert :ok = UmyaSpreadsheet.set_ole_object_requires(ole_object, "xl")
      assert {:ok, "xl"} = UmyaSpreadsheet.get_ole_object_requires(ole_object)

      # Test program ID
      assert :ok = UmyaSpreadsheet.set_ole_object_prog_id(ole_object, "Excel.Sheet.12")
      assert {:ok, "Excel.Sheet.12"} = UmyaSpreadsheet.get_ole_object_prog_id(ole_object)

      # Test extension
      assert :ok = UmyaSpreadsheet.set_ole_object_extension(ole_object, "xlsx")
      assert {:ok, "xlsx"} = UmyaSpreadsheet.get_ole_object_extension(ole_object)
    end

    test "can work with embedded object property values" do
      {:ok, properties} = UmyaSpreadsheet.new_embedded_object_properties()

      # Test embedded object program ID
      assert :ok = UmyaSpreadsheet.set_embedded_object_prog_id(properties, "Word.Document.12")
      assert {:ok, "Word.Document.12"} = UmyaSpreadsheet.get_embedded_object_prog_id(properties)

      # Test shape ID
      assert :ok = UmyaSpreadsheet.set_embedded_object_shape_id(properties, 123)
      assert {:ok, 123} = UmyaSpreadsheet.get_embedded_object_shape_id(properties)
    end
  end

  describe "OLE Object data management" do
    test "can set and get object data" do
      {:ok, ole_object} = UmyaSpreadsheet.new_ole_object()
      test_data = :binary.bin_to_list("test binary data")

      # Set data - now this should work since the NIF is implemented
      assert :ok = UmyaSpreadsheet.set_ole_object_data(ole_object, test_data)

      # Get data back - now this should work since the NIF is implemented
      assert {:ok, retrieved_data} = UmyaSpreadsheet.get_ole_object_data(ole_object)
      assert is_list(retrieved_data)
      assert :binary.list_to_bin(retrieved_data) == "test binary data"
    end

    test "can check format types" do
      {:ok, ole_object} = UmyaSpreadsheet.new_ole_object()

      # Should have default format checks
      assert {:ok, is_binary} = UmyaSpreadsheet.is_ole_object_binary_format(ole_object)
      assert is_boolean(is_binary)

      assert {:ok, is_excel} = UmyaSpreadsheet.is_ole_object_excel_format(ole_object)
      assert is_boolean(is_excel)
    end
  end

  describe "Worksheet integration" do
    test "can set and get OLE objects from worksheet" do
      # Create spreadsheet and OLE objects
      {:ok, spreadsheet} = UmyaSpreadsheet.new()
      {:ok, ole_objects} = UmyaSpreadsheet.new_ole_objects()
      {:ok, ole_object} = UmyaSpreadsheet.new_ole_object()

      # Add object to collection
      assert :ok = UmyaSpreadsheet.add_ole_object(ole_objects, ole_object)

      # Set OLE objects to worksheet
      assert :ok =
               UmyaSpreadsheet.set_ole_objects_to_worksheet(spreadsheet, "Sheet1", ole_objects)

      # Get OLE objects from worksheet
      assert {:ok, retrieved_objects} =
               UmyaSpreadsheet.get_ole_objects_from_worksheet(spreadsheet, "Sheet1")

      # Should have our object
      assert {:ok, 1} = UmyaSpreadsheet.get_ole_objects_count(retrieved_objects)
    end
  end

  describe "File operations" do
    test "can create OLE object with data" do
      test_data = :binary.bin_to_list("test file content")

      # Create with data - now this should work since the NIF is implemented
      assert {:ok, ole_object} = UmyaSpreadsheet.new_ole_object_with_data(test_data, "docx")

      # Should have the data
      assert {:ok, retrieved_data} = UmyaSpreadsheet.get_ole_object_data(ole_object)
      assert is_list(retrieved_data)
      assert :binary.list_to_bin(retrieved_data) == "test file content"

      # Should have correct properties
      assert {:ok, "Word.Document.12"} = UmyaSpreadsheet.get_ole_object_prog_id(ole_object)
      assert {:ok, "docx"} = UmyaSpreadsheet.get_ole_object_extension(ole_object)
    end

    test "can load and save OLE objects from files" do
      # Create a temporary file
      test_dir = "test/result_files"
      File.mkdir_p!(test_dir)
      test_file_path = "#{test_dir}/ole_test_temp.txt"
      output_file_path = "#{test_dir}/ole_test_output.txt"

      # Write test content to the file
      test_content = "Test OLE content"
      File.write!(test_file_path, test_content)

      # Load file into OLE object
      assert {:ok, ole_object} = UmyaSpreadsheet.new_ole_object_from_file(test_file_path)

      # Save OLE content to another file
      assert :ok = UmyaSpreadsheet.save_ole_object_to_file(ole_object, output_file_path)

      # Verify the saved content
      assert File.read!(output_file_path) == test_content

      # Clean up
      File.rm!(test_file_path)
      File.rm!(output_file_path)
    end
  end

  describe "Helper functions" do
    test "can determine ProgID from extension" do
      assert OleObjects.determine_prog_id("docx") == "Word.Document.12"
      assert OleObjects.determine_prog_id("xlsx") == "Excel.Sheet.12"
      assert OleObjects.determine_prog_id("pptx") == "PowerPoint.Show.12"
      assert OleObjects.determine_prog_id("pdf") == "AcroExch.Document"
      assert OleObjects.determine_prog_id("txt") == "txtfile"
      assert OleObjects.determine_prog_id("unknown") == "Package"
    end
  end
end
