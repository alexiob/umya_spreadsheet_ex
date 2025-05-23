defmodule UmyaSpreadsheet.AlignmentTest do
  use ExUnit.Case, async: true

  @output_path "test/result_files/alignment_test.xlsx"

  setup_all do
    # Make sure the result files directory exists
    File.mkdir_p!("test/result_files")
    :ok
  end

  test "set cell alignment" do
    # Create a new spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Set some values
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Left Top")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Center Top")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C1", "Right Top")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "Left Center")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", "Center Center")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C2", "Right Center")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "Left Bottom")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B3", "Center Bottom")
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C3", "Right Bottom")

    # Make cells larger for better visibility of alignment
    :ok = UmyaSpreadsheet.set_row_height(spreadsheet, "Sheet1", 1, 30.0)
    :ok = UmyaSpreadsheet.set_row_height(spreadsheet, "Sheet1", 2, 30.0)
    :ok = UmyaSpreadsheet.set_row_height(spreadsheet, "Sheet1", 3, 30.0)
    :ok = UmyaSpreadsheet.set_column_width(spreadsheet, "Sheet1", "A", 20.0)
    :ok = UmyaSpreadsheet.set_column_width(spreadsheet, "Sheet1", "B", 20.0)
    :ok = UmyaSpreadsheet.set_column_width(spreadsheet, "Sheet1", "C", 20.0)

    # Set different alignments
    # Row 1 - Top alignments
    :ok = UmyaSpreadsheet.set_cell_alignment(spreadsheet, "Sheet1", "A1", "left", "top")
    :ok = UmyaSpreadsheet.set_cell_alignment(spreadsheet, "Sheet1", "B1", "center", "top")
    :ok = UmyaSpreadsheet.set_cell_alignment(spreadsheet, "Sheet1", "C1", "right", "top")

    # Row 2 - Center vertical alignments
    :ok = UmyaSpreadsheet.set_cell_alignment(spreadsheet, "Sheet1", "A2", "left", "center")
    :ok = UmyaSpreadsheet.set_cell_alignment(spreadsheet, "Sheet1", "B2", "center", "center")
    :ok = UmyaSpreadsheet.set_cell_alignment(spreadsheet, "Sheet1", "C2", "right", "center")

    # Row 3 - Bottom alignments
    :ok = UmyaSpreadsheet.set_cell_alignment(spreadsheet, "Sheet1", "A3", "left", "bottom")
    :ok = UmyaSpreadsheet.set_cell_alignment(spreadsheet, "Sheet1", "B3", "center", "bottom")
    :ok = UmyaSpreadsheet.set_cell_alignment(spreadsheet, "Sheet1", "C3", "right", "bottom")

    # Add some background color to better see the cell boundaries
    for cell <- ["A1", "B1", "C1", "A2", "B2", "C2", "A3", "B3", "C3"] do
      :ok = UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", cell, "#EEEEEE")
    end

    # Write the file
    :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)

    # Verify the file was created
    assert File.exists?(@output_path)
  end

  test "set alignment with justify and distributed" do
    # Create a new spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Set some longer text values
    long_text =
      "This is a longer text that should be displayed in multiple lines with different alignment settings"

    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", long_text)
    :ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", long_text)

    # Make cells larger but constrained width for word wrapping
    :ok = UmyaSpreadsheet.set_row_height(spreadsheet, "Sheet1", 1, 60.0)
    :ok = UmyaSpreadsheet.set_row_height(spreadsheet, "Sheet1", 2, 60.0)
    :ok = UmyaSpreadsheet.set_column_width(spreadsheet, "Sheet1", "A", 30.0)

    # Enable text wrapping
    :ok = UmyaSpreadsheet.set_wrap_text(spreadsheet, "Sheet1", "A1", true)
    :ok = UmyaSpreadsheet.set_wrap_text(spreadsheet, "Sheet1", "A2", true)

    # Set different alignments
    :ok = UmyaSpreadsheet.set_cell_alignment(spreadsheet, "Sheet1", "A1", "justify", "center")

    :ok =
      UmyaSpreadsheet.set_cell_alignment(
        spreadsheet,
        "Sheet1",
        "A2",
        "distributed",
        "distributed"
      )

    # Add some background color to better see the cell boundaries
    :ok = UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "A1", "#EEFFEE")
    :ok = UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "A2", "#EEEEFF")

    # Write the file
    :ok = UmyaSpreadsheet.write(spreadsheet, @output_path)

    # Verify the file was created
    assert File.exists?(@output_path)
  end
end
