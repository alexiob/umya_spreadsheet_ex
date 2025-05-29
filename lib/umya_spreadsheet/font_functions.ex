defmodule UmyaSpreadsheet.FontFunctions do
  @moduledoc """
  Functions for manipulating font styles in a spreadsheet.
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaSpreadsheet.ErrorHandling
  alias UmyaNative

  @doc """
  Sets the font color for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `color` - The color code (e.g., "#FF0000" for red)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      # Set text to red
      :ok = UmyaSpreadsheet.FontFunctions.set_font_color(spreadsheet, "Sheet1", "A1", "#FF0000")
  """
  def set_font_color(%Spreadsheet{reference: ref}, sheet_name, cell_address, color) do
    UmyaNative.set_font_color(ref, sheet_name, cell_address, color)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the font size for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `size` - The font size in points

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.FontFunctions.set_font_size(spreadsheet, "Sheet1", "A1", 14)
  """
  def set_font_size(%Spreadsheet{reference: ref}, sheet_name, cell_address, size) do
    UmyaNative.set_font_size(ref, sheet_name, cell_address, size)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the font weight (bold) for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `is_bold` - Boolean indicating whether the font should be bold

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.FontFunctions.set_font_bold(spreadsheet, "Sheet1", "A1", true)
  """
  def set_font_bold(%Spreadsheet{reference: ref}, sheet_name, cell_address, is_bold) do
    UmyaNative.set_font_bold(ref, sheet_name, cell_address, is_bold)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the font name for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `font_name` - The name of the font

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.FontFunctions.set_font_name(spreadsheet, "Sheet1", "A1", "Arial")
  """
  def set_font_name(%Spreadsheet{reference: ref}, sheet_name, cell_address, font_name) do
    UmyaNative.set_font_name(ref, sheet_name, cell_address, font_name)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the font style to italic for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `is_italic` - Boolean indicating whether the font should be italic

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.FontFunctions.set_font_italic(spreadsheet, "Sheet1", "A1", true)
  """
  def set_font_italic(%Spreadsheet{reference: ref}, sheet_name, cell_address, is_italic) do
    UmyaNative.set_font_italic(ref, sheet_name, cell_address, is_italic)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the font underline style for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `underline_style` - The underline style ("single", "double", "none")

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.FontFunctions.set_font_underline(spreadsheet, "Sheet1", "A1", "single")
  """
  def set_font_underline(%Spreadsheet{reference: ref}, sheet_name, cell_address, underline_style) do
    UmyaNative.set_font_underline(ref, sheet_name, cell_address, underline_style)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the font strikethrough property for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `is_strikethrough` - Boolean indicating whether the font should have strikethrough

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.FontFunctions.set_font_strikethrough(spreadsheet, "Sheet1", "A1", true)
  """
  def set_font_strikethrough(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_address,
        is_strikethrough
      ) do
    UmyaNative.set_font_strikethrough(ref, sheet_name, cell_address, is_strikethrough)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the font family for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `font_family` - The font family: "roman" (serif), "swiss" (sans-serif), "modern" (monospace),
    "script", or "decorative"

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.FontFunctions.set_font_family(spreadsheet, "Sheet1", "A1", "swiss")
  """
  def set_font_family(%Spreadsheet{reference: ref}, sheet_name, cell_address, font_family) do
    UmyaNative.set_font_family(ref, sheet_name, cell_address, font_family)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the font scheme for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")
  - `font_scheme` - The font scheme: "major", "minor", or "none"

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.FontFunctions.set_font_scheme(spreadsheet, "Sheet1", "A1", "major")
  """
  def set_font_scheme(%Spreadsheet{reference: ref}, sheet_name, cell_address, font_scheme) do
    UmyaNative.set_font_scheme(ref, sheet_name, cell_address, font_scheme)
    |> ErrorHandling.standardize_result()
  end

  # Font getter functions

  @doc """
  Gets the font name for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - `{:ok, font_name}` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, font_name} = UmyaSpreadsheet.FontFunctions.get_font_name(spreadsheet, "Sheet1", "A1")
  """
  def get_font_name(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    case UmyaNative.get_font_name(ref, sheet_name, cell_address) do
      {:error, _} = error -> error
      font_name -> {:ok, font_name}
    end
  end

  @doc """
  Gets the font size for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - `{:ok, font_size}` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, font_size} = UmyaSpreadsheet.FontFunctions.get_font_size(spreadsheet, "Sheet1", "A1")
  """
  def get_font_size(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    case UmyaNative.get_font_size(ref, sheet_name, cell_address) do
      {:error, _} = error -> error
      font_size -> {:ok, font_size}
    end
  end

  @doc """
  Gets the bold state for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - `{:ok, is_bold}` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, is_bold} = UmyaSpreadsheet.FontFunctions.get_font_bold(spreadsheet, "Sheet1", "A1")
  """
  def get_font_bold(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    case UmyaNative.get_font_bold(ref, sheet_name, cell_address) do
      {:error, _} = error -> error
      is_bold -> {:ok, is_bold}
    end
  end

  @doc """
  Gets the italic state for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - `{:ok, is_italic}` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, is_italic} = UmyaSpreadsheet.FontFunctions.get_font_italic(spreadsheet, "Sheet1", "A1")
  """
  def get_font_italic(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    case UmyaNative.get_font_italic(ref, sheet_name, cell_address) do
      {:error, _} = error -> error
      is_italic -> {:ok, is_italic}
    end
  end

  @doc """
  Gets the underline style for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - `{:ok, underline_style}` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, underline_style} = UmyaSpreadsheet.FontFunctions.get_font_underline(spreadsheet, "Sheet1", "A1")
  """
  def get_font_underline(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    case UmyaNative.get_font_underline(ref, sheet_name, cell_address) do
      {:error, _} = error -> error
      underline_style -> {:ok, underline_style}
    end
  end

  @doc """
  Gets the strikethrough state for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - `{:ok, is_strikethrough}` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, is_strikethrough} = UmyaSpreadsheet.FontFunctions.get_font_strikethrough(spreadsheet, "Sheet1", "A1")
  """
  def get_font_strikethrough(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    case UmyaNative.get_font_strikethrough(ref, sheet_name, cell_address) do
      {:error, _} = error -> error
      is_strikethrough -> {:ok, is_strikethrough}
    end
  end

  @doc """
  Gets the font family for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - `{:ok, font_family}` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, font_family} = UmyaSpreadsheet.FontFunctions.get_font_family(spreadsheet, "Sheet1", "A1")
  """
  def get_font_family(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    case UmyaNative.get_font_family(ref, sheet_name, cell_address) do
      {:error, _} = error -> error
      font_family -> {:ok, font_family}
    end
  end

  @doc """
  Gets the font scheme for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - `{:ok, font_scheme}` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, font_scheme} = UmyaSpreadsheet.FontFunctions.get_font_scheme(spreadsheet, "Sheet1", "A1")
  """
  def get_font_scheme(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    case UmyaNative.get_font_scheme(ref, sheet_name, cell_address) do
      {:error, _} = error -> error
      font_scheme -> {:ok, font_scheme}
    end
  end

  @doc """
  Gets the font color for a cell.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `sheet_name` - The name of the sheet
  - `cell_address` - The cell address (e.g., "A1", "B5")

  ## Returns

  - `{:ok, font_color}` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, font_color} = UmyaSpreadsheet.FontFunctions.get_font_color(spreadsheet, "Sheet1", "A1")
  """
  def get_font_color(%Spreadsheet{reference: ref}, sheet_name, cell_address) do
    case UmyaNative.get_font_color(ref, sheet_name, cell_address) do
      {:error, _} = error -> error
      font_color -> {:ok, font_color}
    end
  end
end
