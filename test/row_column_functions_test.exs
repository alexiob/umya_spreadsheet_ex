defmodule UmyaSpreadsheet.RowColumnFunctionsTest do
  use ExUnit.Case, async: true
  alias UmyaSpreadsheet.RowColumnFunctions

  @output_path "test/result_files/row_column_functions_test_output.xlsx"

  setup do
    {:ok, spreadsheet} = UmyaSpreadsheet.new()
    File.rm(@output_path)
    %{spreadsheet: spreadsheet}
  end

  describe "column width functions" do
    test "set and get column width", %{spreadsheet: spreadsheet} do
      # Set column width
      assert :ok = RowColumnFunctions.set_column_width(spreadsheet, "Sheet1", "A", 20.5)
      
      # Get column width
      assert {:ok, width} = RowColumnFunctions.get_column_width(spreadsheet, "Sheet1", "A")
      assert width == 20.5
      
      # Test different column
      assert :ok = RowColumnFunctions.set_column_width(spreadsheet, "Sheet1", "B", 15.0)
      assert {:ok, 15.0} = RowColumnFunctions.get_column_width(spreadsheet, "Sheet1", "B")
      
      # Test default width for unset column
      assert {:ok, default_width} = RowColumnFunctions.get_column_width(spreadsheet, "Sheet1", "Z")
      assert default_width == 8.43  # Default column width
    end

    test "set and get column auto width", %{spreadsheet: spreadsheet} do
      # Set column auto width
      assert :ok = RowColumnFunctions.set_column_auto_width(spreadsheet, "Sheet1", "A", true)
      
      # Get column auto width
      assert {:ok, auto_width} = RowColumnFunctions.get_column_auto_width(spreadsheet, "Sheet1", "A")
      assert auto_width == true
      
      # Set to false
      assert :ok = RowColumnFunctions.set_column_auto_width(spreadsheet, "Sheet1", "A", false)
      assert {:ok, false} = RowColumnFunctions.get_column_auto_width(spreadsheet, "Sheet1", "A")
      
      # Test default auto width for unset column
      assert {:ok, default_auto} = RowColumnFunctions.get_column_auto_width(spreadsheet, "Sheet1", "Z")
      assert default_auto == false  # Default auto width is false
    end

    test "get column hidden status", %{spreadsheet: spreadsheet} do
      # Test default hidden status for column
      assert {:ok, hidden} = RowColumnFunctions.get_column_hidden(spreadsheet, "Sheet1", "A")
      assert hidden == false  # Default hidden is false
      
      # Test another column
      assert {:ok, false} = RowColumnFunctions.get_column_hidden(spreadsheet, "Sheet1", "B")
    end
  end

  describe "row height functions" do
    test "set and get row height", %{spreadsheet: spreadsheet} do
      # Set row height
      assert :ok = RowColumnFunctions.set_row_height(spreadsheet, "Sheet1", 1, 25.0)
      
      # Get row height
      assert {:ok, height} = RowColumnFunctions.get_row_height(spreadsheet, "Sheet1", 1)
      assert height == 25.0
      
      # Test different row
      assert :ok = RowColumnFunctions.set_row_height(spreadsheet, "Sheet1", 2, 30.5)
      assert {:ok, 30.5} = RowColumnFunctions.get_row_height(spreadsheet, "Sheet1", 2)
      
      # Test default height for unset row
      assert {:ok, default_height} = RowColumnFunctions.get_row_height(spreadsheet, "Sheet1", 100)
      assert default_height == 15.0  # Default row height
    end

    test "get row hidden status", %{spreadsheet: spreadsheet} do
      # Test default hidden status for row
      assert {:ok, hidden} = RowColumnFunctions.get_row_hidden(spreadsheet, "Sheet1", 1)
      assert hidden == false  # Default hidden is false
      
      # Test another row
      assert {:ok, false} = RowColumnFunctions.get_row_hidden(spreadsheet, "Sheet1", 5)
    end

    test "set row style", %{spreadsheet: spreadsheet} do
      # Set row style with background and font colors
      assert :ok = RowColumnFunctions.set_row_style(spreadsheet, "Sheet1", 1, "#FF0000", "#FFFFFF")
      
      # This should not error, but we can't easily test the visual result without more complex assertions
      # The test mainly verifies the function doesn't crash
    end
  end

  describe "copy styling functions" do
    test "copy row styling", %{spreadsheet: spreadsheet} do
      # Set up some initial styling
      assert :ok = RowColumnFunctions.set_row_height(spreadsheet, "Sheet1", 1, 25.0)
      assert :ok = RowColumnFunctions.set_row_style(spreadsheet, "Sheet1", 1, "#FF0000", "#FFFFFF")
      
      # Copy styling from row 1 to row 2
      assert :ok = RowColumnFunctions.copy_row_styling(spreadsheet, "Sheet1", 1, 2)
      
      # Copy styling with specific column range
      assert :ok = RowColumnFunctions.copy_row_styling(spreadsheet, "Sheet1", 1, 3, 1, 5)
    end

    test "copy column styling", %{spreadsheet: spreadsheet} do
      # Set up some initial styling
      assert :ok = RowColumnFunctions.set_column_width(spreadsheet, "Sheet1", "A", 20.0)
      assert :ok = RowColumnFunctions.set_column_auto_width(spreadsheet, "Sheet1", "A", true)
      
      # Copy styling from column A (index 1) to column B (index 2)
      assert :ok = RowColumnFunctions.copy_column_styling(spreadsheet, "Sheet1", 1, 2)
      
      # Copy styling with specific row range
      assert :ok = RowColumnFunctions.copy_column_styling(spreadsheet, "Sheet1", 1, 3, 1, 10)
    end
  end

  describe "error handling" do
    test "operations on non-existent sheet", %{spreadsheet: spreadsheet} do
      # Test getter functions with non-existent sheet
      assert {:error, "Sheet not found"} = RowColumnFunctions.get_column_width(spreadsheet, "NonExistent", "A")
      assert {:error, "Sheet not found"} = RowColumnFunctions.get_column_auto_width(spreadsheet, "NonExistent", "A")
      assert {:error, "Sheet not found"} = RowColumnFunctions.get_column_hidden(spreadsheet, "NonExistent", "A")
      assert {:error, "Sheet not found"} = RowColumnFunctions.get_row_height(spreadsheet, "NonExistent", 1)
      assert {:error, "Sheet not found"} = RowColumnFunctions.get_row_hidden(spreadsheet, "NonExistent", 1)
      
      # Test setter functions with non-existent sheet
      assert {:error, "Sheet not found"} = RowColumnFunctions.set_column_width(spreadsheet, "NonExistent", "A", 20.0)
      assert {:error, "Sheet not found"} = RowColumnFunctions.set_column_auto_width(spreadsheet, "NonExistent", "A", true)
      assert {:error, "Sheet not found"} = RowColumnFunctions.set_row_height(spreadsheet, "NonExistent", 1, 20.0)
      assert {:error, "Sheet not found"} = RowColumnFunctions.set_row_style(spreadsheet, "NonExistent", 1, "#FF0000", "#FFFFFF")
    end
  end

  describe "integration with main module" do
    test "functions are accessible via main UmyaSpreadsheet module", %{spreadsheet: spreadsheet} do
      # Test that the new getter functions are properly delegated
      assert :ok = UmyaSpreadsheet.set_column_width(spreadsheet, "Sheet1", "A", 18.0)
      assert {:ok, 18.0} = UmyaSpreadsheet.get_column_width(spreadsheet, "Sheet1", "A")
      
      assert :ok = UmyaSpreadsheet.set_column_auto_width(spreadsheet, "Sheet1", "B", true)
      assert {:ok, true} = UmyaSpreadsheet.get_column_auto_width(spreadsheet, "Sheet1", "B")
      
      assert {:ok, false} = UmyaSpreadsheet.get_column_hidden(spreadsheet, "Sheet1", "C")
      
      assert :ok = UmyaSpreadsheet.set_row_height(spreadsheet, "Sheet1", 1, 22.0)
      assert {:ok, 22.0} = UmyaSpreadsheet.get_row_height(spreadsheet, "Sheet1", 1)
      
      assert {:ok, false} = UmyaSpreadsheet.get_row_hidden(spreadsheet, "Sheet1", 2)
    end
  end

  describe "file persistence" do
    test "column and row settings persist after saving and loading", %{spreadsheet: spreadsheet} do
      # Set various properties
      assert :ok = RowColumnFunctions.set_column_width(spreadsheet, "Sheet1", "A", 25.0)
      assert :ok = RowColumnFunctions.set_column_auto_width(spreadsheet, "Sheet1", "B", true)
      assert :ok = RowColumnFunctions.set_row_height(spreadsheet, "Sheet1", 1, 30.0)
      
      # Add some data to make the file meaningful
      assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Test Data")
      assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "More Data")
      
      # Save the file
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)
      
      # Load the file
      {:ok, loaded_spreadsheet} = UmyaSpreadsheet.read(@output_path)
      
      # Verify the properties persist
      assert {:ok, 25.0} = RowColumnFunctions.get_column_width(loaded_spreadsheet, "Sheet1", "A")
      assert {:ok, true} = RowColumnFunctions.get_column_auto_width(loaded_spreadsheet, "Sheet1", "B")
      assert {:ok, 30.0} = RowColumnFunctions.get_row_height(loaded_spreadsheet, "Sheet1", 1)
      
      # Verify data persists too
      assert {:ok, "Test Data"} = UmyaSpreadsheet.get_cell_value(loaded_spreadsheet, "Sheet1", "A1")
      assert {:ok, "More Data"} = UmyaSpreadsheet.get_cell_value(loaded_spreadsheet, "Sheet1", "B1")
    end
  end
end
