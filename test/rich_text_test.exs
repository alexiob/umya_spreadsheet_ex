defmodule RichTextTest do
  use ExUnit.Case
  doctest UmyaSpreadsheet

  setup do
    # Create a temporary file
    temp_path = System.tmp_dir!() |> Path.join("rich_text_test.xlsx")

    on_exit(fn ->
      if File.exists?(temp_path) do
        File.rm!(temp_path)
      end
    end)

    {:ok, temp_path: temp_path}
  end

  test "create and manipulate rich text", %{temp_path: temp_path} do
    # Create a new spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Create rich text
    rich_text = UmyaSpreadsheet.RichText.create()

    # Create font properties for bold text
    bold_props = %{
      "bold" => true,
      "name" => "Arial",
      "size" => "12",
      "color" => "FF0000"
    }

    # Create font properties for normal text
    normal_props = %{
      "bold" => false,
      "name" => "Arial",
      "size" => "12",
      "color" => "000000"
    }

    # Add formatted text to rich text
    :ok = UmyaSpreadsheet.RichText.add_formatted_text(rich_text, "Bold Text", bold_props)
    :ok = UmyaSpreadsheet.RichText.add_formatted_text(rich_text, " and Normal Text", normal_props)

    # Set rich text in a cell
    :ok = UmyaSpreadsheet.RichText.set_cell_rich_text(spreadsheet, "Sheet1", "A1", rich_text)

    # Get rich text from cell
    retrieved_rich_text = UmyaSpreadsheet.RichText.get_cell_rich_text(spreadsheet, "Sheet1", "A1")

    # Get plain text from rich text
    plain_text = UmyaSpreadsheet.RichText.get_plain_text(retrieved_rich_text)
    assert plain_text == "Bold Text and Normal Text"

    # Get elements from rich text
    elements = UmyaSpreadsheet.RichText.get_elements(retrieved_rich_text)
    assert length(elements) == 2

    # Test first element (bold text) - Skip element text test since function is disabled
    [first_element | _] = elements
    first_text = UmyaSpreadsheet.RichText.get_element_text(first_element)
    assert first_text == "Bold Text"

    {:ok, first_props} =
      UmyaSpreadsheet.RichText.get_element_font_properties(first_element)

    assert first_props.bold == "true"

    # Save the file
    :ok = UmyaSpreadsheet.write(spreadsheet, temp_path)

    # Verify file was created
    assert File.exists?(temp_path)
  end

  test "create rich text from HTML", %{temp_path: temp_path} do
    # Create a new spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Create rich text from HTML
    html = "<b>Bold Text</b> and <i>Italic Text</i>"
    rich_text = UmyaSpreadsheet.RichText.create_from_html(html)

    # Set rich text in a cell
    :ok = UmyaSpreadsheet.RichText.set_cell_rich_text(spreadsheet, "Sheet1", "B1", rich_text)

    # Get plain text
    plain_text = UmyaSpreadsheet.RichText.get_plain_text(rich_text)
    assert plain_text == "Bold Text and Italic Text"

    # Convert back to HTML
    output_html = UmyaSpreadsheet.RichText.to_html(rich_text)
    assert String.contains?(output_html, "Bold Text")
    assert String.contains?(output_html, "Italic Text")

    # Save the file
    :ok = UmyaSpreadsheet.write(spreadsheet, temp_path)

    # Verify file was created
    assert File.exists?(temp_path)
  end

  test "create text elements manually", %{temp_path: temp_path} do
    # Create a new spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Create rich text
    rich_text = UmyaSpreadsheet.RichText.create()

    # Create text elements
    bold_props = %{
      "bold" => true,
      "name" => "Arial",
      "size" => "14",
      "color" => "FF0000"
    }

    text_element = UmyaSpreadsheet.RichText.create_text_element("Bold Red Text", bold_props)
    :ok = UmyaSpreadsheet.RichText.add_text_element(rich_text, text_element)

    # Set rich text in a cell
    :ok = UmyaSpreadsheet.RichText.set_cell_rich_text(spreadsheet, "Sheet1", "C1", rich_text)

    # Verify the text
    retrieved_rich_text = UmyaSpreadsheet.RichText.get_cell_rich_text(spreadsheet, "Sheet1", "C1")
    plain_text = UmyaSpreadsheet.RichText.get_plain_text(retrieved_rich_text)
    assert plain_text == "Bold Red Text"

    # Save the file
    :ok = UmyaSpreadsheet.write(spreadsheet, temp_path)

    # Verify file was created
    assert File.exists?(temp_path)
  end
end
