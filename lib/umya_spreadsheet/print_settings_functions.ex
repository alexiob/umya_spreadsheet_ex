defmodule UmyaSpreadsheet.PrintSettings do
  @moduledoc """
  Functions for configuring print settings in worksheets.
  """
  
  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaNative

  @doc """
  Sets the page orientation for a specific sheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `orientation` - The orientation type: "portrait" or "landscape"

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.set_page_orientation(spreadsheet, "Sheet1", "landscape")
  """
  def set_page_orientation(%Spreadsheet{reference: ref}, sheet_name, orientation) do
    case UmyaNative.set_page_orientation(ref, sheet_name, orientation) do
      {:ok, :ok} -> :ok
      result -> result
    end
  end

  @doc """
  Sets the paper size for a specific sheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `paper_size` - The paper size code (e.g., 9 for A4)

  ## Paper Size Codes
     1 = Letter (8.5 x 11 in)
     5 = Legal (8.5 x 14 in)
     9 = A4 (210 x 297 mm)
     8 = A3 (297 x 420 mm)
     7 = A5 (148 x 210 mm)

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.set_paper_size(spreadsheet, "Sheet1", 9)
  """
  def set_paper_size(%Spreadsheet{reference: ref}, sheet_name, paper_size) do
    case UmyaNative.set_paper_size(ref, sheet_name, paper_size) do
      {:ok, :ok} -> :ok
      result -> result
    end
  end

  @doc """
  Sets the page scale percentage for a specific sheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `scale` - The scale percentage (e.g., 100 for 100%, 80 for 80%)

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.set_page_scale(spreadsheet, "Sheet1", 80)
  """
  def set_page_scale(%Spreadsheet{reference: ref}, sheet_name, scale) do
    case UmyaNative.set_page_scale(ref, sheet_name, scale) do
      {:ok, :ok} -> :ok
      result -> result
    end
  end

  @doc """
  Sets the fit-to-page options for a specific sheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `width` - The number of pages wide to fit the content
  - `height` - The number of pages tall to fit the content

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.set_fit_to_page(spreadsheet, "Sheet1", 1, 2)
  """
  def set_fit_to_page(%Spreadsheet{reference: ref}, sheet_name, width, height) do
    case UmyaNative.set_fit_to_page(ref, sheet_name, width, height) do
      {:ok, :ok} -> :ok
      result -> result
    end
  end

  @doc """
  Sets the page margins for a specific sheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `top` - The top margin in inches
  - `right` - The right margin in inches
  - `bottom` - The bottom margin in inches
  - `left` - The left margin in inches

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.set_page_margins(spreadsheet, "Sheet1", 1.0, 0.75, 1.0, 0.75)
  """
  def set_page_margins(%Spreadsheet{reference: ref}, sheet_name, top, right, bottom, left) do
    case UmyaNative.set_page_margins(ref, sheet_name, top, right, bottom, left) do
      {:ok, :ok} -> :ok
      result -> result
    end
  end

  @doc """
  Sets the header and footer margins for a specific sheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `header` - The header margin in inches
  - `footer` - The footer margin in inches

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.set_header_footer_margins(spreadsheet, "Sheet1", 0.5, 0.5)
  """
  def set_header_footer_margins(%Spreadsheet{reference: ref}, sheet_name, header, footer) do
    case UmyaNative.set_header_footer_margins(ref, sheet_name, header, footer) do
      {:ok, :ok} -> :ok
      result -> result
    end
  end

  @doc """
  Sets the header text for a specific sheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `header` - The header text, which can include special formatting codes

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.set_header(spreadsheet, "Sheet1", "&C&\\"Arial,Bold\\"Confidential")
  """
  def set_header(%Spreadsheet{reference: ref}, sheet_name, header) do
    case UmyaNative.set_header(ref, sheet_name, header) do
      {:ok, :ok} -> :ok
      result -> result
    end
  end

  @doc """
  Sets the footer text for a specific sheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `footer` - The footer text, which can include special formatting codes

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.set_footer(spreadsheet, "Sheet1", "&RPage &P of &N")
  """
  def set_footer(%Spreadsheet{reference: ref}, sheet_name, footer) do
    case UmyaNative.set_footer(ref, sheet_name, footer) do
      {:ok, :ok} -> :ok
      result -> result
    end
  end

  @doc """
  Sets whether to center the print horizontally and/or vertically.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `horizontal` - Whether to center horizontally
  - `vertical` - Whether to center vertically

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.set_print_centered(spreadsheet, "Sheet1", true, false)
  """
  def set_print_centered(%Spreadsheet{reference: ref}, sheet_name, horizontal, vertical) do
    case UmyaNative.set_print_centered(ref, sheet_name, horizontal, vertical) do
      {:ok, :ok} -> :ok
      result -> result
    end
  end

  @doc """
  Sets the print area for a specific sheet.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `print_area` - The range to be printed (e.g., "A1:H20")

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.set_print_area(spreadsheet, "Sheet1", "A1:H20")
  """
  def set_print_area(%Spreadsheet{reference: ref}, sheet_name, print_area) do
    case UmyaNative.set_print_area(ref, sheet_name, print_area) do
      {:ok, :ok} -> :ok
      result -> result
    end
  end

  @doc """
  Sets the rows and columns to repeat on each printed page (print titles).

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `rows` - The rows to repeat (e.g., "1:2")
  - `columns` - The columns to repeat (e.g., "A:B")

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.set_print_titles(spreadsheet, "Sheet1", "1:2", "A:B")
  """
  def set_print_titles(%Spreadsheet{reference: ref}, sheet_name, rows, columns) do
    case UmyaNative.set_print_titles(ref, sheet_name, rows, columns) do
      {:ok, :ok} -> :ok
      result -> result
    end
  end
end
