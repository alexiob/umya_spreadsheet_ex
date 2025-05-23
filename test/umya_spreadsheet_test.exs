defmodule UmyaSpreadsheetTest do
  use ExUnit.Case, async: true
  doctest UmyaSpreadsheet

  @temp_file "test/result_files/test_output.xlsx"

  setup do
    # Ensure our test directory exists
    File.mkdir_p!("test/result_files")
  end

  describe "file operations" do
    test "new spreadsheet" do
      assert {:ok, spreadsheet} = UmyaSpreadsheet.new()
      assert is_struct(spreadsheet, UmyaSpreadsheet.Spreadsheet)
      assert ["Sheet1"] = UmyaSpreadsheet.get_sheet_names(spreadsheet)
    end

    test "create and write a spreadsheet" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Write a value to the spreadsheet
      assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Hello")
      assert {:ok, "Hello"} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "A1")

      # Write the spreadsheet to disk
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @temp_file)
      assert File.exists?(@temp_file)
    end

    test "read a spreadsheet" do
      # First create a file
      {:ok, original} = UmyaSpreadsheet.new()
      assert :ok = UmyaSpreadsheet.set_cell_value(original, "Sheet1", "A1", "Hello World")
      assert :ok = UmyaSpreadsheet.write(original, @temp_file)

      # Now read it back
      assert {:ok, loaded} = UmyaSpreadsheet.read(@temp_file)
      assert {:ok, "Hello World"} = UmyaSpreadsheet.get_cell_value(loaded, "Sheet1", "A1")
    end

    test "lazy read a spreadsheet" do
      # First create a file
      {:ok, original} = UmyaSpreadsheet.new()
      assert :ok = UmyaSpreadsheet.set_cell_value(original, "Sheet1", "A1", "Lazy Load Test")
      assert :ok = UmyaSpreadsheet.add_sheet(original, "Sheet2")
      assert :ok = UmyaSpreadsheet.set_cell_value(original, "Sheet2", "B2", "Second Sheet")
      assert :ok = UmyaSpreadsheet.write(original, @temp_file)

      # Now lazy read it
      assert {:ok, loaded} = UmyaSpreadsheet.lazy_read(@temp_file)
      assert ["Sheet1", "Sheet2"] = UmyaSpreadsheet.get_sheet_names(loaded)
      assert {:ok, "Lazy Load Test"} = UmyaSpreadsheet.get_cell_value(loaded, "Sheet1", "A1")
      assert {:ok, "Second Sheet"} = UmyaSpreadsheet.get_cell_value(loaded, "Sheet2", "B2")
    end
  end

  describe "sheet operations" do
    test "add sheet" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()
      assert :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "NewSheet")
      assert ["Sheet1", "NewSheet"] = UmyaSpreadsheet.get_sheet_names(spreadsheet)
    end

    test "get sheet names" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()
      assert :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "Sheet2")
      assert :ok = UmyaSpreadsheet.add_sheet(spreadsheet, "Sheet3")

      sheet_names = UmyaSpreadsheet.get_sheet_names(spreadsheet)
      assert length(sheet_names) == 3
      assert "Sheet1" in sheet_names
      assert "Sheet2" in sheet_names
      assert "Sheet3" in sheet_names
    end
  end

  describe "cell operations" do
    test "get and set cell value" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Test with string
      assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Test String")
      assert {:ok, "Test String"} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "A1")

      # Test with number
      assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", 123)
      assert {:ok, "123"} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "B2")

      # Test with atom
      assert :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C3", :atom_value)
      assert {:ok, "atom_value"} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "C3")
    end

    test "move range" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Setup test data
      :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Item 1")
      :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "Item 2")
      :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "Item 3")

      # Move the range
      assert :ok = UmyaSpreadsheet.move_range(spreadsheet, "Sheet1", "A1:A3", 0, 1)

      # Verify data is moved
      assert {:ok, "Item 1"} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "B1")
      assert {:ok, "Item 2"} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "B2")
      assert {:ok, "Item 3"} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "B3")

      # Original cells should be empty now
      assert {:ok, ""} = UmyaSpreadsheet.get_cell_value(spreadsheet, "Sheet1", "A1")
    end
  end

  describe "style operations" do
    test "set background color" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      # Set background color using named color
      assert :ok = UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "A1", "red")
      assert :ok = UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "A2", "blue")

      # Set background color using hex value
      assert :ok = UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "A3", "#00FF00")

      # Write to file to ensure changes are applied
      assert :ok = UmyaSpreadsheet.write(spreadsheet, @temp_file)
      assert File.exists?(@temp_file)
    end
  end

  describe "error handling" do
    test "get cell from non-existent sheet" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      assert {:error, :not_found} =
               UmyaSpreadsheet.get_cell_value(spreadsheet, "NonExistentSheet", "A1")
    end

    test "set cell in non-existent sheet" do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()

      assert {:error, :not_found} =
               UmyaSpreadsheet.set_cell_value(spreadsheet, "NonExistentSheet", "A1", "Test")
    end

    test "read non-existent file" do
      assert {:error, _} = UmyaSpreadsheet.read("non_existent_file.xlsx")
    end
  end
end
