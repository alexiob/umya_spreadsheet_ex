# Print Settings and Page Setup

This guide explains how to configure print settings and page setup options in UmyaSpreadsheet.

## Overview

UmyaSpreadsheet provides various functions to control how your worksheets appear when printed. These include:

- Page orientation (portrait/landscape)
- Paper size
- Page scaling
- Fit-to options
- Page margins
- Headers and footers
- Print centering
- Print area
- Print titles (repeating rows/columns)

## Basic Print Setup

### Page Orientation

Set the page orientation to portrait or landscape:

```elixir
# Set to landscape orientation
UmyaSpreadsheet.set_page_orientation(spreadsheet, "Sheet1", "landscape")

# Set to portrait orientation
UmyaSpreadsheet.set_page_orientation(spreadsheet, "Sheet1", "portrait")
```

### Paper Size

Set the paper size using standard size codes:

```elixir
# Set to A4 paper (code 9)
UmyaSpreadsheet.set_paper_size(spreadsheet, "Sheet1", 9)

# Set to US Letter (code 1)
UmyaSpreadsheet.set_paper_size(spreadsheet, "Sheet1", 1)
```

Common paper size codes:

- 1: Letter (8.5 x 11 in)
- 5: Legal (8.5 x 14 in)
- 9: A4 (210 x 297 mm)
- 10: A3 (297 x 420 mm)

### Page Scaling

Scale the content to a percentage of its original size:

```elixir
# Scale to 75%
UmyaSpreadsheet.set_page_scale(spreadsheet, "Sheet1", 75)
```

Valid scale values range from 10 to 400 percent.

### Fit-to Options

Make your content fit a specific number of pages:

```elixir
# Fit to 1 page wide by 1 page tall
UmyaSpreadsheet.set_fit_to_page(spreadsheet, "Sheet1", 1, 1)

# Fit to 1 page wide by however many pages needed vertically (0 means no limit)
UmyaSpreadsheet.set_fit_to_page(spreadsheet, "Sheet1", 1, 0)
```

## Page Margins

Set the page margins in inches:

```elixir
# Set margins (top, right, bottom, left)
UmyaSpreadsheet.set_page_margins(spreadsheet, "Sheet1", 1.0, 0.75, 1.0, 0.75)
```

Set header and footer margins specifically:

```elixir
# Set header margin to 0.5 inches and footer margin to 0.5 inches
UmyaSpreadsheet.set_header_footer_margins(spreadsheet, "Sheet1", 0.5, 0.5)
```

## Headers and Footers

Set header content with formatting codes:

```elixir
# Set a centered, bold header
UmyaSpreadsheet.set_header(spreadsheet, "Sheet1", "&C&\"Arial,Bold\"Confidential Document")
```

Set footer content:

```elixir
# Set right-aligned page numbers
UmyaSpreadsheet.set_footer(spreadsheet, "Sheet1", "&RPage &P of &N")
```

### Header/Footer Formatting Codes

- Position codes:
  - `&L`: Left section
  - `&C`: Center section
  - `&R`: Right section

- Font formatting:
  - `&"Font,Style"`: Set font name and style
  - `&B`: Bold
  - `&I`: Italic
  - `&U`: Underline
  - `&S`: Strikethrough
  - `&nn`: Font size (where nn is the size)

- Content codes:
  - `&P`: Current page number
  - `&N`: Total pages
  - `&D`: Current date
  - `&T`: Current time
  - `&F`: File name
  - `&A`: Sheet name

## Print Centering

Center the content on the page:

```elixir
# Center horizontally only
UmyaSpreadsheet.set_print_centered(spreadsheet, "Sheet1", true, false)

# Center both horizontally and vertically
UmyaSpreadsheet.set_print_centered(spreadsheet, "Sheet1", true, true)
```

## Print Area

Define a specific area to print:

```elixir
# Only print cells A1 through H20
UmyaSpreadsheet.set_print_area(spreadsheet, "Sheet1", "A1:H20")
```

## Print Titles

Set rows and/or columns to repeat on each printed page:

```elixir
# Repeat rows 1 and 2 at the top of each page
UmyaSpreadsheet.set_print_titles(spreadsheet, "Sheet1", "1:2", "")

# Repeat columns A and B on each page
UmyaSpreadsheet.set_print_titles(spreadsheet, "Sheet1", "", "A:B")

# Repeat both rows and columns
UmyaSpreadsheet.set_print_titles(spreadsheet, "Sheet1", "1:2", "A:B")
```

## Complete Example

Here's a complete example of setting up a worksheet for printing:

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Configure print settings
UmyaSpreadsheet.set_page_orientation(spreadsheet, "Sheet1", "landscape")
UmyaSpreadsheet.set_paper_size(spreadsheet, "Sheet1", 9)  # A4
UmyaSpreadsheet.set_fit_to_page(spreadsheet, "Sheet1", 1, 0)  # Fit to 1 page wide
UmyaSpreadsheet.set_page_margins(spreadsheet, "Sheet1", 0.75, 0.75, 0.75, 0.75)
UmyaSpreadsheet.set_header(spreadsheet, "Sheet1", "&L&D&R&F")
UmyaSpreadsheet.set_footer(spreadsheet, "Sheet1", "&RPage &P of &N")
UmyaSpreadsheet.set_print_area(spreadsheet, "Sheet1", "A1:G20")
UmyaSpreadsheet.set_print_titles(spreadsheet, "Sheet1", "1:1", "")

# Write the file
UmyaSpreadsheet.write(spreadsheet, "print_settings_example.xlsx")
```
