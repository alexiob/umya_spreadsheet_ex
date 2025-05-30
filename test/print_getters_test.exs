defmodule UmyaSpreadsheet.PrintGettersTest do
  use ExUnit.Case, async: true

  describe "print settings getter functions" do
    setup do
      {:ok, spreadsheet} = UmyaSpreadsheet.new()
      %{spreadsheet: spreadsheet}
    end

    test "get_page_orientation returns default portrait", %{spreadsheet: spreadsheet} do
      orientation = UmyaSpreadsheet.get_page_orientation(spreadsheet, "Sheet1")
      assert orientation == "portrait"
    end

    test "get_page_orientation returns landscape after setting", %{spreadsheet: spreadsheet} do
      :ok = UmyaSpreadsheet.set_page_orientation(spreadsheet, "Sheet1", "landscape")
      orientation = UmyaSpreadsheet.get_page_orientation(spreadsheet, "Sheet1")
      assert orientation == "landscape"
    end

    test "get_paper_size returns default value", %{spreadsheet: spreadsheet} do
      paper_size = UmyaSpreadsheet.get_paper_size(spreadsheet, "Sheet1")
      assert is_integer(paper_size)
    end

    test "get_paper_size returns set value", %{spreadsheet: spreadsheet} do
      # A4
      :ok = UmyaSpreadsheet.set_paper_size(spreadsheet, "Sheet1", 9)
      paper_size = UmyaSpreadsheet.get_paper_size(spreadsheet, "Sheet1")
      assert paper_size == 9
    end

    test "get_page_scale returns default value", %{spreadsheet: spreadsheet} do
      scale = UmyaSpreadsheet.get_page_scale(spreadsheet, "Sheet1")
      assert is_integer(scale)
    end

    test "get_page_scale returns set value", %{spreadsheet: spreadsheet} do
      :ok = UmyaSpreadsheet.set_page_scale(spreadsheet, "Sheet1", 75)
      scale = UmyaSpreadsheet.get_page_scale(spreadsheet, "Sheet1")
      assert scale == 75
    end

    test "get_fit_to_page returns default values", %{spreadsheet: spreadsheet} do
      {width, height} = UmyaSpreadsheet.get_fit_to_page(spreadsheet, "Sheet1")
      assert is_integer(width) and is_integer(height)
    end

    test "get_fit_to_page returns set values", %{spreadsheet: spreadsheet} do
      :ok = UmyaSpreadsheet.set_fit_to_page(spreadsheet, "Sheet1", 1, 2)
      {width, height} = UmyaSpreadsheet.get_fit_to_page(spreadsheet, "Sheet1")
      assert width == 1 and height == 2
    end

    test "get_page_margins returns default values", %{spreadsheet: spreadsheet} do
      {top, right, bottom, left} = UmyaSpreadsheet.get_page_margins(spreadsheet, "Sheet1")
      assert is_float(top) and is_float(right) and is_float(bottom) and is_float(left)
    end

    test "get_page_margins returns set values", %{spreadsheet: spreadsheet} do
      :ok = UmyaSpreadsheet.set_page_margins(spreadsheet, "Sheet1", 1.0, 0.75, 1.0, 0.75)
      {top, right, bottom, left} = UmyaSpreadsheet.get_page_margins(spreadsheet, "Sheet1")
      assert top == 1.0 and right == 0.75 and bottom == 1.0 and left == 0.75
    end

    test "get_header_footer_margins returns default values", %{spreadsheet: spreadsheet} do
      {header, footer} = UmyaSpreadsheet.get_header_footer_margins(spreadsheet, "Sheet1")
      assert is_float(header) and is_float(footer)
    end

    test "get_header_footer_margins returns set values", %{spreadsheet: spreadsheet} do
      :ok = UmyaSpreadsheet.set_header_footer_margins(spreadsheet, "Sheet1", 0.5, 0.5)
      {header, footer} = UmyaSpreadsheet.get_header_footer_margins(spreadsheet, "Sheet1")
      assert header == 0.5 and footer == 0.5
    end

    test "get_header returns string", %{spreadsheet: spreadsheet} do
      header = UmyaSpreadsheet.get_header(spreadsheet, "Sheet1")
      assert is_binary(header)
    end

    test "get_footer returns string", %{spreadsheet: spreadsheet} do
      footer = UmyaSpreadsheet.get_footer(spreadsheet, "Sheet1")
      assert is_binary(footer)
    end

    test "get_print_centered returns default false values", %{spreadsheet: spreadsheet} do
      {horizontal, vertical} = UmyaSpreadsheet.get_print_centered(spreadsheet, "Sheet1")
      assert is_boolean(horizontal) and is_boolean(vertical)
    end

    test "get_print_centered returns set values", %{spreadsheet: spreadsheet} do
      :ok = UmyaSpreadsheet.set_print_centered(spreadsheet, "Sheet1", true, false)
      {horizontal, vertical} = UmyaSpreadsheet.get_print_centered(spreadsheet, "Sheet1")
      assert horizontal == true and vertical == false
    end

    test "get_print_area returns empty string by default", %{spreadsheet: spreadsheet} do
      print_area = UmyaSpreadsheet.get_print_area(spreadsheet, "Sheet1")
      assert is_binary(print_area)
    end

    test "get_print_titles returns empty strings by default", %{spreadsheet: spreadsheet} do
      {rows, cols} = UmyaSpreadsheet.get_print_titles(spreadsheet, "Sheet1")
      assert is_binary(rows) and is_binary(cols)
    end

    test "getter functions return error for non-existent sheet", %{spreadsheet: spreadsheet} do
      assert {:error, _} = UmyaSpreadsheet.get_page_orientation(spreadsheet, "NonExistent")
      assert {:error, _} = UmyaSpreadsheet.get_paper_size(spreadsheet, "NonExistent")
      assert {:error, _} = UmyaSpreadsheet.get_page_scale(spreadsheet, "NonExistent")
      assert {:error, _} = UmyaSpreadsheet.get_fit_to_page(spreadsheet, "NonExistent")
      assert {:error, _} = UmyaSpreadsheet.get_page_margins(spreadsheet, "NonExistent")
      assert {:error, _} = UmyaSpreadsheet.get_header_footer_margins(spreadsheet, "NonExistent")
      assert {:error, _} = UmyaSpreadsheet.get_header(spreadsheet, "NonExistent")
      assert {:error, _} = UmyaSpreadsheet.get_footer(spreadsheet, "NonExistent")
      assert {:error, _} = UmyaSpreadsheet.get_print_centered(spreadsheet, "NonExistent")
      assert {:error, _} = UmyaSpreadsheet.get_print_area(spreadsheet, "NonExistent")
      assert {:error, _} = UmyaSpreadsheet.get_print_titles(spreadsheet, "NonExistent")
    end
  end
end
