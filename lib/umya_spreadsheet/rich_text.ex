defmodule UmyaSpreadsheet.RichText do
  @moduledoc """
  Rich Text functionality for UmyaSpreadsheet.

  This module provides support for formatted text within cells, allowing multiple
  font styles, colors, and formatting within a single cell.

  ## Examples

      # Create a simple rich text
      rich_text = UmyaSpreadsheet.RichText.create()

      # Add formatted text
      UmyaSpreadsheet.RichText.add_formatted_text(rich_text, "Bold text", %{bold: true})
      UmyaSpreadsheet.RichText.add_formatted_text(rich_text, " and normal text", %{})

      # Set rich text to a cell
      UmyaSpreadsheet.RichText.set_cell_rich_text(spreadsheet, "Sheet1", "A1", rich_text)

      # Create rich text from HTML
      rich_text = UmyaSpreadsheet.RichText.create_from_html("<b>Bold</b> and <i>italic</i> text")
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaSpreadsheet.ErrorHandling
  alias UmyaNative

  # Helper function to convert atom keys to string keys for Rust compatibility
  defp normalize_font_props(font_props) when is_map(font_props) do
    Enum.into(font_props, %{}, fn
      {key, value} when is_atom(key) -> {Atom.to_string(key), value}
      {key, value} -> {key, value}
    end)
  end

  # Helper function to convert string keys back to atom keys for Elixir interface
  defp denormalize_font_props(font_props) when is_map(font_props) do
    Enum.into(font_props, %{}, fn
      {key, value} when is_binary(key) ->
        atom_key =
          case key do
            "bold" -> :bold
            "italic" -> :italic
            "underline" -> :underline
            "strikethrough" -> :strikethrough
            "size" -> :size
            "name" -> :name
            "color" -> :color
            _ -> String.to_atom(key)
          end

        {atom_key, value}

      {key, value} ->
        {key, value}
    end)
  end

  @doc """
  Creates a new empty RichText object.

  Returns a Rich Text resource that can be used to build formatted text.

  ## Examples

      iex> rich_text = UmyaSpreadsheet.RichText.create()
      iex> is_reference(rich_text)
      true
  """
  def create do
    UmyaNative.create_rich_text()
  end

  @doc """
  Creates a RichText object from an HTML string.

  Parses HTML to create formatted rich text with font styling.

  ## Parameters

  - `html` - HTML string to parse

  ## Examples

      iex> html = "<b>Bold</b> and <i>italic</i> text"
      iex> rich_text = UmyaSpreadsheet.RichText.create_from_html(html)
      iex> is_reference(rich_text)
      true
  """
  def create_from_html(html) when is_binary(html) do
    UmyaNative.create_rich_text_from_html(html)
  end

  @doc """
  Creates a TextElement with text and optional font properties.

  A TextElement represents a portion of text with consistent formatting.

  ## Parameters

  - `text` - The text content
  - `font_props` - Map of font properties (optional)

  ## Font Properties

  - `:bold` - Boolean for bold text
  - `:italic` - Boolean for italic text
  - `:underline` - Boolean for underlined text
  - `:strikethrough` - Boolean for strikethrough text
  - `:size` - Font size (number)
  - `:name` - Font name (string)
  - `:color` - Color (hex string like "#FF0000" or color name)

  ## Examples

      iex> element = UmyaSpreadsheet.RichText.create_text_element("Bold text", %{bold: true, color: "#FF0000"})
      iex> is_reference(element)
      true
  """
  def create_text_element(text, font_props \\ %{}) when is_binary(text) and is_map(font_props) do
    UmyaNative.create_text_element(text, normalize_font_props(font_props))
  end

  @doc """
  Adds a TextElement to a RichText object.

  ## Parameters

  - `rich_text` - RichText resource
  - `text_element` - TextElement resource to add

  ## Examples

      iex> rich_text = UmyaSpreadsheet.RichText.create()
      iex> element = UmyaSpreadsheet.RichText.create_text_element("Test", %{bold: true})
      iex> UmyaSpreadsheet.RichText.add_text_element(rich_text, element)
      :ok
  """
  def add_text_element(rich_text, text_element) do
    UmyaNative.add_text_element_to_rich_text(rich_text, text_element)
  end

  @doc """
  Adds formatted text directly to a RichText object.

  This is a convenience function that creates a TextElement and adds it in one step.

  ## Parameters

  - `rich_text` - RichText resource
  - `text` - Text content to add
  - `font_props` - Map of font properties (optional)

  ## Examples

      iex> rich_text = UmyaSpreadsheet.RichText.create()
      iex> UmyaSpreadsheet.RichText.add_formatted_text(rich_text, "Bold text", %{bold: true})
      :ok
      iex> UmyaSpreadsheet.RichText.add_formatted_text(rich_text, " normal text", %{})
      :ok
  """
  def add_formatted_text(rich_text, text, font_props \\ %{})
      when is_binary(text) and is_map(font_props) do
    UmyaNative.add_formatted_text_to_rich_text(rich_text, text, normalize_font_props(font_props))
  end

  @doc """
  Sets rich text to a specific cell.

  ## Parameters

  - `spreadsheet` - UmyaSpreadsheet resource
  - `sheet_name` - Name of the worksheet
  - `coordinate` - Cell coordinate (e.g., "A1")
  - `rich_text` - RichText resource to set

  ## Examples

      iex> spreadsheet = UmyaSpreadsheet.new_file()
      iex> rich_text = UmyaSpreadsheet.RichText.create()
      iex> UmyaSpreadsheet.RichText.add_formatted_text(rich_text, "Test", %{bold: true})
      iex> UmyaSpreadsheet.RichText.set_cell_rich_text(spreadsheet, "Sheet1", "A1", rich_text)
      :ok
  """
  def set_cell_rich_text(%Spreadsheet{reference: ref}, sheet_name, coordinate, rich_text)
      when is_binary(sheet_name) and is_binary(coordinate) do
    UmyaNative.set_cell_rich_text(ref, sheet_name, coordinate, rich_text)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets rich text from a specific cell.

  ## Parameters

  - `spreadsheet` - UmyaSpreadsheet resource
  - `sheet_name` - Name of the worksheet
  - `coordinate` - Cell coordinate (e.g., "A1")

  ## Returns

  Returns a RichText resource. If the cell doesn't contain rich text, returns an empty RichText.

  ## Examples

      iex> spreadsheet = UmyaSpreadsheet.new_file()
      iex> rich_text = UmyaSpreadsheet.RichText.get_cell_rich_text(spreadsheet, "Sheet1", "A1")
      iex> is_reference(rich_text)
      true
  """
  def get_cell_rich_text(%Spreadsheet{reference: ref}, sheet_name, coordinate)
      when is_binary(sheet_name) and is_binary(coordinate) do
    UmyaNative.get_cell_rich_text(ref, sheet_name, coordinate)
  end

  @doc """
  Gets plain text from a RichText object.

  Extracts all text content without formatting.

  ## Parameters

  - `rich_text` - RichText resource

  ## Examples

      iex> rich_text = UmyaSpreadsheet.RichText.create()
      iex> UmyaSpreadsheet.RichText.add_formatted_text(rich_text, "Test", %{bold: true})
      iex> UmyaSpreadsheet.RichText.get_plain_text(rich_text)
      "Test"
  """
  def get_plain_text(rich_text) do
    UmyaNative.get_rich_text_plain_text(rich_text)
  end

  @doc """
  Converts RichText to HTML representation.

  ## Parameters

  - `rich_text` - RichText resource

  ## Examples

      iex> rich_text = UmyaSpreadsheet.RichText.create()
      iex> UmyaSpreadsheet.RichText.add_formatted_text(rich_text, "Bold", %{bold: true})
      iex> UmyaSpreadsheet.RichText.to_html(rich_text)
      "<b>Bold</b>"
  """
  def to_html(rich_text) do
    UmyaNative.rich_text_to_html(rich_text)
  end

  @doc """
  Gets all text elements from a RichText object.

  ## Parameters

  - `rich_text` - RichText resource

  ## Returns

  List of TextElement resources.

  ## Examples

      iex> rich_text = UmyaSpreadsheet.RichText.create()
      iex> UmyaSpreadsheet.RichText.add_formatted_text(rich_text, "Test", %{bold: true})
      iex> elements = UmyaSpreadsheet.RichText.get_elements(rich_text)
      iex> is_list(elements)
      true
  """
  def get_elements(rich_text) do
    UmyaNative.get_rich_text_elements(rich_text)
  end

  @doc """
  Gets text content from a TextElement.

  ## Parameters

  - `text_element` - TextElement resource

  ## Examples

      iex> element = UmyaSpreadsheet.RichText.create_text_element("Test", %{})
      iex> UmyaSpreadsheet.RichText.get_element_text(element)
      "Test"
  """
  def get_element_text(text_element) do
    UmyaNative.get_text_element_text(text_element)
  end

  @doc """
  Gets font properties from a TextElement.

  ## Parameters

  - `text_element` - TextElement resource

  ## Returns

  `{:ok, %{}}` - Map of font properties including name, size, bold, italic, strikethrough, underline, and color.

  ## Examples

      iex> element = UmyaSpreadsheet.RichText.create_text_element("Test", %{bold: true, size: 12})
      iex> {:ok, props} = UmyaSpreadsheet.RichText.get_element_font_properties(element)
      iex> props[:bold]
      "true"
  """
  def get_element_font_properties(text_element) do
    case UmyaNative.get_text_element_font_properties(text_element) do
      %{} = props -> {:ok, denormalize_font_props(props)}
      {:error, _} = error -> error
    end
  end
end
