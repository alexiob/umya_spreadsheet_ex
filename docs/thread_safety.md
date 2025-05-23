# Thread Safety and Concurrent Operations

This guide provides information about using UmyaSpreadsheet in multi-threaded environments and concurrent operations.

## Overview

UmyaSpreadsheet provides the ability to work with Excel files in concurrent operations. Understanding the thread safety characteristics of the library is important when developing applications that need to handle multiple spreadsheets simultaneously or operate on the same spreadsheet from multiple threads.

## Thread Safety Guidelines

When working with UmyaSpreadsheet in a multi-threaded environment, keep these guidelines in mind:

### What Works Well in Concurrent Scenarios

1. **Creating Multiple Independent Spreadsheets**: You can safely create and manipulate multiple
   spreadsheet objects in parallel threads without conflicts.

2. **Drawing Operations on Different Spreadsheets**: Adding shapes, text boxes, and connectors
   to different spreadsheet objects in parallel is safe.

3. **Reading From Multiple Threads**: Reading operations (getting cell values, etc.) from different
   threads on the same spreadsheet are generally safe.

### Potential Issues in Concurrent Scenarios

1. **Race Conditions with Cell Value Updates**: When multiple threads read a cell value and then
   update it based on that value, race conditions can occur. The operation isn't atomic, so the
   final value may not reflect all updates.

2. **Spreadsheet Modifications**: Modifying the same spreadsheet from multiple threads without
   synchronization may lead to unexpected results.

### Best Practices for Concurrent Operations

1. **Spreadsheet Per Thread**: For maximum thread safety, create and manipulate separate spreadsheet
   objects in each thread, then combine the results if needed.

2. **Add Synchronization**: If you need to modify the same spreadsheet from multiple threads,
   consider adding synchronization in your application code (e.g., using `Task.await` or process
   synchronization mechanisms).

3. **Atomic Operations**: For operations like incrementing a cell value, consider using a dedicated
   process to handle the operation atomically.

## Example: Concurrent Spreadsheet Creation

```elixir
# Create multiple spreadsheets concurrently
tasks = for i <- 1..5 do
  Task.async(fn ->
    # Each task creates its own spreadsheet
    {:ok, spreadsheet} = UmyaSpreadsheet.new()

    # Perform operations on the spreadsheet
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Spreadsheet #{i}")
    UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Created concurrently")

    # Save the spreadsheet to a unique file
    file_path = "spreadsheet_#{i}.xlsx"
    UmyaSpreadsheet.write(spreadsheet, file_path)

    # Return the path and spreadsheet reference
    {file_path, spreadsheet}
  end)
end

# Wait for all tasks to complete
results = Enum.map(tasks, &Task.await(&1, 10000))

# Process the results
Enum.each(results, fn {path, _spreadsheet} ->
  IO.puts("Created: #{path}")
end)
```

## Example: Synchronized Access to a Shared Spreadsheet

```elixir
# Create an Agent to manage access to a shared spreadsheet
{:ok, agent} = Agent.start_link(fn ->
  {:ok, spreadsheet} = UmyaSpreadsheet.new()
  spreadsheet
end)

# Function to safely increment a cell value
defp increment_cell(agent, sheet, cell, amount) do
  Agent.update(agent, fn spreadsheet ->
    # Get current value
    {:ok, current_value} = UmyaSpreadsheet.get_cell_value(spreadsheet, sheet, cell)
    current_value = if is_number(current_value), do: current_value, else: 0

    # Set new value
    UmyaSpreadsheet.set_cell_value(spreadsheet, sheet, cell, current_value + amount)

    # Return the spreadsheet reference (required for Agent.update)
    spreadsheet
  end)
end

# Create multiple tasks that safely update the same spreadsheet
tasks = for i <- 1..10 do
  Task.async(fn ->
    # Each task increments the counter
    increment_cell(agent, "Sheet1", "A1", 1)

    # Add a record of the operation
    Agent.update(agent, fn spreadsheet ->
      UmyaSpreadsheet.set_cell_value(
        spreadsheet,
        "Sheet1",
        "B#{i+1}",
        "Increment #{i} at #{DateTime.utc_now()}"
      )
      spreadsheet
    end)
  end)
end

# Wait for all tasks to complete
Enum.each(tasks, &Task.await(&1, 10000))

# Save the final spreadsheet
Agent.get(agent, fn spreadsheet ->
  UmyaSpreadsheet.write(spreadsheet, "synchronized_spreadsheet.xlsx")
end)

# Stop the agent
Agent.stop(agent)
```

## Best Practices Summary

1. **Use Independent Spreadsheets**: When possible, create separate spreadsheet instances for each thread.

2. **Employ Process Synchronization**: Use Elixir's built-in concurrency primitives (`Agent`, `GenServer`, etc.) to manage access to shared spreadsheets.

3. **Batch Operations**: Instead of many small concurrent operations, consider batching changes and applying them sequentially.

4. **Test Thoroughly**: Always test your concurrent code patterns with the specific operations you'll be performing.

5. **Document Thread Safety Assumptions**: Make sure to document which operations are thread-safe in your application.

## See Also

- [Shapes and Drawing](shapes_and_drawing.html) - For shape drawing operations that can be done concurrently
- [Sheet Operations](sheet_operations.html) - For managing worksheets in concurrent environments
