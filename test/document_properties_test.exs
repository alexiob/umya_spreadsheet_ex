defmodule UmyaSpreadsheet.DocumentPropertiesTest do
  use ExUnit.Case
  doctest UmyaSpreadsheet.DocumentProperties

  alias UmyaSpreadsheet.DocumentProperties

  @output_file_path "test/result_files/document_properties_test_output.xlsx"

  setup do
    # Ensure result directory exists
    File.mkdir_p!("test/result_files")

    # Create a new spreadsheet for testing
    {:ok, spreadsheet} = UmyaSpreadsheet.new()
    UmyaSpreadsheet.CellFunctions.set_cell_value(spreadsheet, "Sheet1", "A1", "Test Data")

    on_exit(fn ->
      # Clean up output files if they exist
      if File.exists?(@output_file_path), do: File.rm!(@output_file_path)
    end)

    {:ok, spreadsheet: spreadsheet}
  end

  describe "custom properties" do
    test "set and get string custom property", %{spreadsheet: spreadsheet} do
      property_name = "ProjectName"
      property_value = "Project Alpha"

      # Set the property
      assert :ok =
               DocumentProperties.set_custom_property_string(
                 spreadsheet,
                 property_name,
                 property_value
               )

      # Get the property
      assert {:ok, ^property_value} =
               DocumentProperties.get_custom_property(spreadsheet, property_name)
    end

    test "set and get number custom property", %{spreadsheet: spreadsheet} do
      property_name = "Version"
      property_value = 1.5

      # Set the property
      assert :ok =
               DocumentProperties.set_custom_property_number(
                 spreadsheet,
                 property_name,
                 property_value
               )

      # Get the property
      assert {:ok, ^property_value} =
               DocumentProperties.get_custom_property(spreadsheet, property_name)
    end

    test "set and get integer custom property", %{spreadsheet: spreadsheet} do
      property_name = "Count"
      property_value = 42

      # Set the property
      assert :ok =
               DocumentProperties.set_custom_property_number(
                 spreadsheet,
                 property_name,
                 property_value
               )

      # Get the property
      assert {:ok, ^property_value} =
               DocumentProperties.get_custom_property(spreadsheet, property_name)
    end

    test "set and get boolean custom property", %{spreadsheet: spreadsheet} do
      property_name = "IsApproved"
      property_value = true

      # Set the property
      assert :ok =
               DocumentProperties.set_custom_property_bool(
                 spreadsheet,
                 property_name,
                 property_value
               )

      # Get the property
      assert {:ok, ^property_value} =
               DocumentProperties.get_custom_property(spreadsheet, property_name)

      # Test false value
      property_value2 = false

      assert :ok =
               DocumentProperties.set_custom_property_bool(
                 spreadsheet,
                 "IsPublic",
                 property_value2
               )

      assert {:ok, ^property_value2} =
               DocumentProperties.get_custom_property(spreadsheet, "IsPublic")
    end

    test "set and get date custom property", %{spreadsheet: spreadsheet} do
      property_name = "ReleaseDate"
      property_value = "2024-12-31T00:00:00Z"

      # Set the property
      assert :ok =
               DocumentProperties.set_custom_property_date(
                 spreadsheet,
                 property_name,
                 property_value
               )

      # Get the property - don't compare exact value as timezone handling may differ
      {:ok, result} = DocumentProperties.get_custom_property(spreadsheet, property_name)
      assert is_binary(result)
      assert String.contains?(result, "2024-12-31")
    end

    test "convenience set_custom_property function with automatic type detection", %{
      spreadsheet: spreadsheet
    } do
      # Test string
      assert :ok =
               DocumentProperties.set_custom_property(spreadsheet, "StringProp", "test string")

      assert {:ok, "test string"} =
               DocumentProperties.get_custom_property(spreadsheet, "StringProp")

      # Test number
      assert :ok = DocumentProperties.set_custom_property(spreadsheet, "NumberProp", 3.14)
      assert {:ok, 3.14} = DocumentProperties.get_custom_property(spreadsheet, "NumberProp")

      # Test boolean
      assert :ok = DocumentProperties.set_custom_property(spreadsheet, "BoolProp", true)
      assert {:ok, true} = DocumentProperties.get_custom_property(spreadsheet, "BoolProp")
    end

    test "get non-existent custom property", %{spreadsheet: spreadsheet} do
      assert {:error, :not_found} =
               DocumentProperties.get_custom_property(spreadsheet, "NonExistent")
    end

    test "remove custom property", %{spreadsheet: spreadsheet} do
      property_name = "TempProperty"
      property_value = "temporary"

      # Set the property
      assert :ok =
               DocumentProperties.set_custom_property_string(
                 spreadsheet,
                 property_name,
                 property_value
               )

      assert {:ok, ^property_value} =
               DocumentProperties.get_custom_property(spreadsheet, property_name)

      # Remove the property
      assert :ok = DocumentProperties.remove_custom_property(spreadsheet, property_name)

      assert {:error, :not_found} =
               DocumentProperties.get_custom_property(spreadsheet, property_name)
    end

    test "remove non-existent custom property", %{spreadsheet: spreadsheet} do
      # May return :ok or {:error, :not_found} - both are valid behaviors
      result = DocumentProperties.remove_custom_property(spreadsheet, "NonExistent")
      assert result == :ok or result == {:error, :not_found}
    end

    test "check if custom property exists", %{spreadsheet: spreadsheet} do
      property_name = "ExistenceTest"

      # Property should not exist initially
      assert {:ok, false} = DocumentProperties.has_custom_property(spreadsheet, property_name)

      # Set the property
      assert :ok =
               DocumentProperties.set_custom_property_string(spreadsheet, property_name, "exists")

      # Property should now exist
      assert {:ok, true} = DocumentProperties.has_custom_property(spreadsheet, property_name)

      # Remove the property
      assert :ok = DocumentProperties.remove_custom_property(spreadsheet, property_name)

      # Property should not exist again
      assert {:ok, false} = DocumentProperties.has_custom_property(spreadsheet, property_name)
    end

    test "get custom property names", %{spreadsheet: spreadsheet} do
      # Initially should be empty
      assert {:ok, []} = DocumentProperties.get_custom_property_names(spreadsheet)

      # Add some properties
      assert :ok = DocumentProperties.set_custom_property_string(spreadsheet, "Prop1", "value1")
      assert :ok = DocumentProperties.set_custom_property_number(spreadsheet, "Prop2", 42)
      assert :ok = DocumentProperties.set_custom_property_bool(spreadsheet, "Prop3", true)

      # Get property names
      assert {:ok, names} = DocumentProperties.get_custom_property_names(spreadsheet)
      assert length(names) == 3
      assert "Prop1" in names
      assert "Prop2" in names
      assert "Prop3" in names
    end

    test "get custom properties count", %{spreadsheet: spreadsheet} do
      # Initially should be 0
      assert {:ok, 0} = DocumentProperties.get_custom_properties_count(spreadsheet)

      # Add some properties
      assert :ok = DocumentProperties.set_custom_property_string(spreadsheet, "Count1", "value1")
      assert {:ok, 1} = DocumentProperties.get_custom_properties_count(spreadsheet)

      assert :ok = DocumentProperties.set_custom_property_number(spreadsheet, "Count2", 42)
      assert {:ok, 2} = DocumentProperties.get_custom_properties_count(spreadsheet)

      # Remove a property
      assert :ok = DocumentProperties.remove_custom_property(spreadsheet, "Count1")
      assert {:ok, 1} = DocumentProperties.get_custom_properties_count(spreadsheet)
    end

    test "clear all custom properties", %{spreadsheet: spreadsheet} do
      # Add some properties
      assert :ok = DocumentProperties.set_custom_property_string(spreadsheet, "Clear1", "value1")
      assert :ok = DocumentProperties.set_custom_property_string(spreadsheet, "Clear2", "value2")
      assert :ok = DocumentProperties.set_custom_property_string(spreadsheet, "Clear3", "value3")

      # Verify they exist
      assert {:ok, 3} = DocumentProperties.get_custom_properties_count(spreadsheet)

      # Clear all properties
      assert :ok = DocumentProperties.clear_custom_properties(spreadsheet)

      # Verify they're gone
      assert {:ok, 0} = DocumentProperties.get_custom_properties_count(spreadsheet)
      assert {:ok, []} = DocumentProperties.get_custom_property_names(spreadsheet)
    end
  end

  describe "core document properties" do
    test "set and get title", %{spreadsheet: spreadsheet} do
      title = "My Test Spreadsheet"
      assert :ok = DocumentProperties.set_title(spreadsheet, title)
      assert {:ok, ^title} = DocumentProperties.get_title(spreadsheet)
    end

    test "set and get description", %{spreadsheet: spreadsheet} do
      description = "This is a test spreadsheet for unit testing"
      assert :ok = DocumentProperties.set_description(spreadsheet, description)
      assert {:ok, ^description} = DocumentProperties.get_description(spreadsheet)
    end

    test "set and get subject", %{spreadsheet: spreadsheet} do
      subject = "Unit Testing"
      assert :ok = DocumentProperties.set_subject(spreadsheet, subject)
      assert {:ok, ^subject} = DocumentProperties.get_subject(spreadsheet)
    end

    test "set and get keywords", %{spreadsheet: spreadsheet} do
      keywords = "test, spreadsheet, elixir, rust"
      assert :ok = DocumentProperties.set_keywords(spreadsheet, keywords)
      assert {:ok, ^keywords} = DocumentProperties.get_keywords(spreadsheet)
    end

    test "set and get creator", %{spreadsheet: spreadsheet} do
      creator = "Test Suite"
      assert :ok = DocumentProperties.set_creator(spreadsheet, creator)
      assert {:ok, ^creator} = DocumentProperties.get_creator(spreadsheet)
    end

    test "set and get last modified by", %{spreadsheet: spreadsheet} do
      last_modified_by = "ExUnit Test Runner"
      assert :ok = DocumentProperties.set_last_modified_by(spreadsheet, last_modified_by)
      assert {:ok, ^last_modified_by} = DocumentProperties.get_last_modified_by(spreadsheet)
    end

    test "set and get category", %{spreadsheet: spreadsheet} do
      category = "Testing"
      assert :ok = DocumentProperties.set_category(spreadsheet, category)
      assert {:ok, ^category} = DocumentProperties.get_category(spreadsheet)
    end

    test "set and get company", %{spreadsheet: spreadsheet} do
      company = "Test Company Inc."
      assert :ok = DocumentProperties.set_company(spreadsheet, company)
      assert {:ok, ^company} = DocumentProperties.get_company(spreadsheet)
    end

    test "set and get manager", %{spreadsheet: spreadsheet} do
      manager = "Test Manager"
      assert :ok = DocumentProperties.set_manager(spreadsheet, manager)
      assert {:ok, ^manager} = DocumentProperties.get_manager(spreadsheet)
    end

    test "set and get created date", %{spreadsheet: spreadsheet} do
      created = "2024-01-01T00:00:00Z"
      assert :ok = DocumentProperties.set_created(spreadsheet, created)
      assert {:ok, ^created} = DocumentProperties.get_created(spreadsheet)
    end

    test "set and get modified date", %{spreadsheet: spreadsheet} do
      modified = "2024-05-28T12:00:00Z"
      assert :ok = DocumentProperties.set_modified(spreadsheet, modified)
      assert {:ok, ^modified} = DocumentProperties.get_modified(spreadsheet)
    end
  end

  describe "convenience functions" do
    test "set multiple properties at once", %{spreadsheet: spreadsheet} do
      properties = %{
        title: "Multi-Set Test",
        description: "Testing multiple property setting",
        creator: "Test Suite",
        company: "Test Corp",
        category: "Unit Test"
      }

      assert :ok = DocumentProperties.set_properties(spreadsheet, properties)

      # Verify all properties were set
      assert {:ok, "Multi-Set Test"} = DocumentProperties.get_title(spreadsheet)

      assert {:ok, "Testing multiple property setting"} =
               DocumentProperties.get_description(spreadsheet)

      assert {:ok, "Test Suite"} = DocumentProperties.get_creator(spreadsheet)
      assert {:ok, "Test Corp"} = DocumentProperties.get_company(spreadsheet)
      assert {:ok, "Unit Test"} = DocumentProperties.get_category(spreadsheet)
    end

    test "set_properties with unknown property returns error", %{spreadsheet: spreadsheet} do
      properties = %{
        title: "Valid Title",
        unknown_property: "Invalid"
      }

      assert {:error, {:unknown_property, :unknown_property}} =
               DocumentProperties.set_properties(spreadsheet, properties)
    end

    test "get all properties", %{spreadsheet: spreadsheet} do
      # Set some properties first
      assert :ok = DocumentProperties.set_title(spreadsheet, "All Props Test")
      assert :ok = DocumentProperties.set_description(spreadsheet, "Testing get all")
      assert :ok = DocumentProperties.set_creator(spreadsheet, "Creator Test")

      # Get all properties
      assert {:ok, all_props} = DocumentProperties.get_all_properties(spreadsheet)

      # Verify the map structure
      assert is_map(all_props)
      assert all_props.title == "All Props Test"
      assert all_props.description == "Testing get all"
      assert all_props.creator == "Creator Test"

      # Should have all expected keys
      expected_keys = [
        :title,
        :description,
        :subject,
        :keywords,
        :creator,
        :last_modified_by,
        :category,
        :company,
        :manager,
        :created,
        :modified
      ]

      for key <- expected_keys do
        assert Map.has_key?(all_props, key)
      end
    end
  end

  describe "persistence" do
    test "properties persist after saving and reloading", %{spreadsheet: spreadsheet} do
      # Set some properties
      assert :ok = DocumentProperties.set_title(spreadsheet, "Persistence Test")
      assert :ok = DocumentProperties.set_description(spreadsheet, "Testing persistence")

      assert :ok =
               DocumentProperties.set_custom_property_string(
                 spreadsheet,
                 "CustomProp",
                 "custom value"
               )

      assert :ok = DocumentProperties.set_custom_property_number(spreadsheet, "CustomNum", 123.45)

      # Save the file
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @output_file_path)

      # Read the file back
      assert {:ok, reloaded_spreadsheet} = UmyaSpreadsheet.read(@output_file_path)

      # Verify core properties persisted
      assert {:ok, "Persistence Test"} = DocumentProperties.get_title(reloaded_spreadsheet)

      assert {:ok, "Testing persistence"} =
               DocumentProperties.get_description(reloaded_spreadsheet)

      # Verify custom properties persisted
      assert {:ok, "custom value"} =
               DocumentProperties.get_custom_property(reloaded_spreadsheet, "CustomProp")

      assert {:ok, 123.45} =
               DocumentProperties.get_custom_property(reloaded_spreadsheet, "CustomNum")
    end
  end

  describe "edge cases" do
    test "empty string properties", %{spreadsheet: spreadsheet} do
      # Test empty strings for core properties
      assert :ok = DocumentProperties.set_title(spreadsheet, "")
      assert {:ok, ""} = DocumentProperties.get_title(spreadsheet)

      # Test empty strings for custom properties
      assert :ok = DocumentProperties.set_custom_property_string(spreadsheet, "EmptyProp", "")
      assert {:ok, ""} = DocumentProperties.get_custom_property(spreadsheet, "EmptyProp")
    end

    test "unicode characters in properties", %{spreadsheet: spreadsheet} do
      unicode_title = "测试标题 🎉 émojis"
      unicode_custom = "Café naïveté résumé 🚀"

      assert :ok = DocumentProperties.set_title(spreadsheet, unicode_title)
      assert {:ok, ^unicode_title} = DocumentProperties.get_title(spreadsheet)

      assert :ok =
               DocumentProperties.set_custom_property_string(
                 spreadsheet,
                 "Unicode",
                 unicode_custom
               )

      assert {:ok, ^unicode_custom} =
               DocumentProperties.get_custom_property(spreadsheet, "Unicode")
    end

    test "very long property values", %{spreadsheet: spreadsheet} do
      # Create a long string (1000 characters)
      long_string = String.duplicate("a", 1000)

      assert :ok = DocumentProperties.set_description(spreadsheet, long_string)
      assert {:ok, ^long_string} = DocumentProperties.get_description(spreadsheet)

      assert :ok =
               DocumentProperties.set_custom_property_string(spreadsheet, "LongProp", long_string)

      assert {:ok, ^long_string} = DocumentProperties.get_custom_property(spreadsheet, "LongProp")
    end

    test "property names with special characters", %{spreadsheet: spreadsheet} do
      special_name = "prop-with.special_chars@123"
      value = "special property"

      assert :ok = DocumentProperties.set_custom_property_string(spreadsheet, special_name, value)
      assert {:ok, ^value} = DocumentProperties.get_custom_property(spreadsheet, special_name)
    end

    test "zero and negative numbers", %{spreadsheet: spreadsheet} do
      assert :ok = DocumentProperties.set_custom_property_number(spreadsheet, "Zero", 0)
      assert {:ok, 0} = DocumentProperties.get_custom_property(spreadsheet, "Zero")

      # For negative floating point values
      assert :ok = DocumentProperties.set_custom_property_number(spreadsheet, "Negative", -42.5)
      {:ok, result} = DocumentProperties.get_custom_property(spreadsheet, "Negative")
      assert_in_delta -42.5, result, 0.0001
    end

    test "very large numbers", %{spreadsheet: spreadsheet} do
      # Close to max float64
      large_number = 1.7976931348623157e+308

      assert :ok =
               DocumentProperties.set_custom_property_number(spreadsheet, "Large", large_number)

      # Large numbers may be returned as strings due to precision limitations
      {:ok, result} = DocumentProperties.get_custom_property(spreadsheet, "Large")

      # Convert to float and compare if it's a string
      result_float = if is_binary(result), do: String.to_float(result), else: result

      # Use a percentage-based threshold for very large numbers
      # 1% difference tolerance
      assert_in_delta large_number, result_float, large_number * 0.01
    end
  end

  describe "integration with existing file" do
    @tag timeout: :infinity
    test "reading and modifying existing file with properties" do
      # Create a temporary test file
      temp_file_path = "test/result_files/temp_test_file.xlsx"

      # Create a new spreadsheet for testing
      {:ok, original_spreadsheet} = UmyaSpreadsheet.new()

      UmyaSpreadsheet.CellFunctions.set_cell_value(
        original_spreadsheet,
        "Sheet1",
        "A1",
        "Test Data"
      )

      # Add some initial properties
      DocumentProperties.set_title(original_spreadsheet, "Original Title")

      DocumentProperties.set_custom_property_string(
        original_spreadsheet,
        "InitialProp",
        "initial value"
      )

      # Save it
      assert :ok = UmyaSpreadsheet.write(original_spreadsheet, temp_file_path)

      # Now read the file back
      {:ok, spreadsheet} = UmyaSpreadsheet.read(temp_file_path)

      # Get existing properties count
      {:ok, initial_count} = DocumentProperties.get_custom_properties_count(spreadsheet)

      # Verify initial properties
      assert {:ok, "Original Title"} = DocumentProperties.get_title(spreadsheet)

      assert {:ok, "initial value"} =
               DocumentProperties.get_custom_property(spreadsheet, "InitialProp")

      # Add some new properties
      assert :ok =
               DocumentProperties.set_custom_property_string(
                 spreadsheet,
                 "AddedProp",
                 "added value"
               )

      assert :ok = DocumentProperties.set_title(spreadsheet, "Modified Title")

      # Verify the property was added
      {:ok, new_count} = DocumentProperties.get_custom_properties_count(spreadsheet)
      assert new_count == initial_count + 1

      assert {:ok, "added value"} =
               DocumentProperties.get_custom_property(spreadsheet, "AddedProp")

      assert {:ok, "Modified Title"} = DocumentProperties.get_title(spreadsheet)

      # Clean up temp file
      on_exit(fn ->
        if File.exists?(temp_file_path), do: File.rm!(temp_file_path)
      end)
    end
  end
end
