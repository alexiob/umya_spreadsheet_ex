# Development Notes

## Future Improvements

1. Add support for more operations:

   - Rename sheet functionality (rename_sheet)
   - Additional conditional formatting options

2. Improve error handling:

   - Rust wrapper should propagate errors to Elixir, and not println
   - Use Elixir's `{:ok, result}` and `{:error, reason}` pattern
   - Handle errors in a more user-friendly way
   - Provide more context in error messages
   - Use Elixir's `with` construct for better error handling
   - More specific error types/messages
   - Better validation for inputs
   - Improved handling of corrupted files

3. Add more comprehensive documentation:
   - Add module docs with overview of architecture
   - Document limitations and compatibility issues
   - Create a troubleshooting guide

## Features Available in Rust but Missing in the Elixir Wrapper

1. File Format Options

The Rust library likely supports more file format options than what's exposed in the Elixir wrapper.
For example, more control over XLSX compression, encryption options, etc.
