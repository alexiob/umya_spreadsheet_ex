# Comments

This guide covers the comment management functionality in UmyaSpreadsheet, allowing you to add, retrieve, update, and remove cell comments in your Excel spreadsheets.

## Adding Comments to Cells

Comments are a great way to provide context, leave notes, or give instructions within your spreadsheet without affecting the cell's value.

```elixir
alias UmyaSpreadsheet

# Create a new spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add some data
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Sales Report")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", 5000)

# Add a comment to cell B2
UmyaSpreadsheet.add_comment(spreadsheet, "Sheet1", "B2", "This value represents Q1 sales.", "Alex")

# Save the spreadsheet
UmyaSpreadsheet.write(spreadsheet, "comments_example.xlsx")
```

The `add_comment` function allows you to specify:
- The spreadsheet to modify
- The sheet name where the cell is located
- The cell address to add the comment to
- The comment text
- The author of the comment (optional)

## Retrieving Comments

You can retrieve the text and author of a comment from a specific cell:

```elixir
# Get comment from cell B2
case UmyaSpreadsheet.get_comment(spreadsheet, "Sheet1", "B2") do
  {:ok, {text, author}} ->
    IO.puts("Comment by #{author}: #{text}")
  {:error, _reason} ->
    IO.puts("No comment found on this cell")
end
```

## Updating Existing Comments

You can update an existing comment on a cell:

```elixir
# Update the comment on cell B2
UmyaSpreadsheet.update_comment(
  spreadsheet,
  "Sheet1",
  "B2",
  "Updated: This value represents Q1 sales in USD.",
  "Alex Smith"  # Updated author name
)
```

If you want to keep the original author, you can omit the author parameter:

```elixir
UmyaSpreadsheet.update_comment(
  spreadsheet,
  "Sheet1",
  "B2",
  "Updated: This value represents Q1 sales in USD."
  # Author remains unchanged
)
```

## Removing Comments

To remove a comment from a cell:

```elixir
UmyaSpreadsheet.remove_comment(spreadsheet, "Sheet1", "B2")
```

## Checking for Comments

You can check if a worksheet contains any comments:

```elixir
if UmyaSpreadsheet.has_comments(spreadsheet, "Sheet1") do
  IO.puts("Sheet1 contains comments")
else
  IO.puts("Sheet1 has no comments")
end
```

## Counting Comments

To get the number of comments in a worksheet:

```elixir
comment_count = UmyaSpreadsheet.get_comments_count(spreadsheet, "Sheet1")
IO.puts("Sheet1 has #{comment_count} comments")
```

## Best Practices

1. **Keep Comments Concise**: Write clear, brief comments that provide valuable context
2. **Include Dates**: For collaborative spreadsheets, consider including dates in your comments
3. **Use Consistently**: Develop a consistent approach to commenting within your organization
4. **Author Attribution**: Always include author information for accountability
5. **Regular Cleanup**: Remove outdated comments to keep spreadsheets clean and focused

## Comment Visibility

When opening a spreadsheet in Excel or other applications:

- Comments are typically indicated by a small triangle in the corner of cells
- Users can hover over cells to see the comments
- Comments retain their author attribution when viewed in Excel
