defmodule TestFormatting do
  def run do
    # Load the library
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Add test data
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "10")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "20")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "30")

    # Add a simple conditional formatting rule
    UmyaSpreadsheet.add_cell_value_rule(
      spreadsheet,
      "Sheet1",
      "A1:A3",
      "greaterThan",
      "15",
      nil,
      # Red
      "#FF0000"
    )

    # Try to retrieve the rules
    rules =
      UmyaSpreadsheet.ConditionalFormatting.get_conditional_formatting_rules(
        spreadsheet,
        "Sheet1"
      )

    IO.inspect(rules, label: "Conditional Formatting Rules")

    # Save the spreadsheet for verification
    UmyaSpreadsheet.write(spreadsheet, "test/result_files/test_cf.xlsx")
  end
end

# Run the test
TestFormatting.run()
