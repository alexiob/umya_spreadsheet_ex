defmodule DebugVmlTest do
  use ExUnit.Case

  @temp_file "test/result_files/debug_vml_simple.xlsx"

  test "debug simple write without VML" do
    # Clean up temp file before test
    if File.exists?(@temp_file), do: File.rm!(@temp_file)

    {:ok, spreadsheet} = UmyaSpreadsheet.new()
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Test")

    # Write without VML
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @temp_file)
    assert File.exists?(@temp_file)

    # Read back
    assert {:ok, _} = UmyaSpreadsheet.read(@temp_file)
  end

  test "debug with VML step by step" do
    # Clean up temp file before test
    if File.exists?(@temp_file), do: File.rm!(@temp_file)

    {:ok, spreadsheet} = UmyaSpreadsheet.new()
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "VML Test")

    # Create VML shape but don't add properties yet
    assert :ok = UmyaSpreadsheet.VmlDrawing.create_shape(spreadsheet, "Sheet1", "shape1")

    # Try to write with just the shape
    assert :ok = UmyaSpreadsheet.write(spreadsheet, @temp_file)
    assert File.exists?(@temp_file)
  end
end
