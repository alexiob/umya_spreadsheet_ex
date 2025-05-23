# API module for data validation functionality
defmodule UmyaSpreadsheet.DataValidation do
  @moduledoc """
  Functions for applying data validation to spreadsheets.

  Data validation allows you to control what users can enter in a cell,
  and provides feedback when invalid data is entered. Thi    # Convert numerical values to strings as required by the Rust NIF
    value1_str = Float.to_string(value1)
    value2_str = if value2, do: Float.to_string(value2), else: nil

    case UmyaNative.add_number_validation(
           UmyaSpreadsheet.unwrap_ref(ref),
           sheet_name,
           cell_range,
           operator,
           value1_str,
           value2_str,
           allow_blank,
           error_title,
           error_message,rovides functions for:

  * List validation (dropdown lists)
  * Number validation (numerical constraints)
  * Date validation (date constraints)
  * Text length validation (character count constraints)
  * Custom formula validation (formula-based constraints)
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaNative

  @doc """
  Adds a dropdown list validation to a range of cells.

  This creates a dropdown list that restricts input to the specified options.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet
    * `cell_range` - Range of cells in A1 notation (e.g., "A1:B10")
    * `list_items` - List of strings to be shown in the dropdown
    * `allow_blank` - Whether to allow blank/empty values (default: true)
    * `error_title` - Optional title for error message when validation fails
    * `error_message` - Optional description for error message when validation fails
    * `prompt_title` - Optional title for input prompt
    * `prompt_message` - Optional description for input prompt

  ## Examples

  ```elixir
  # Add a simple dropdown with fruits
  DataValidation.add_list_validation(
    spreadsheet,
    "Sheet1",
    "A1:A10",
    ["Apple", "Orange", "Banana"],
    true,
    nil,
    nil,
    nil,
    nil
  )

  # Add a dropdown with error message when invalid data is entered
  DataValidation.add_list_validation(
    spreadsheet,
    "Sheet1",
    "B1:B10",
    ["Red", "Green", "Blue"],
    true,
    "Invalid Color",
    "Please select a valid color from the dropdown list",
    "Color Selection",
    "Select a color from the list"
  )
  ```
  """
  @spec add_list_validation(
          Spreadsheet.t(),
          String.t(),
          String.t(),
          [String.t()],
          boolean(),
          String.t() | nil,
          String.t() | nil,
          String.t() | nil,
          String.t() | nil
        ) :: :ok | {:error, atom()}
  def add_list_validation(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_range,
        list_items,
        allow_blank \\ true,
        error_title \\ nil,
        error_message \\ nil,
        prompt_title \\ nil,
        prompt_message \\ nil
      ) do
    case UmyaNative.add_list_validation(
           UmyaSpreadsheet.unwrap_ref(ref),
           sheet_name,
           cell_range,
           list_items,
           allow_blank,
           error_title,
           error_message,
           prompt_title,
           prompt_message
         ) do
      :ok -> :ok
      {:ok, :ok} -> :ok
      {:error, reason} -> {:error, reason}
    end
  end

  @doc """
  Adds number validation to a range of cells.

  This restricts input to numbers that meet specified conditions.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet
    * `cell_range` - Range of cells in A1 notation (e.g., "A1:B10")
    * `operator` - Comparison operator, one of:
      * "between" - Between two values
      * "notBetween" - Not between two values
      * "equal" - Equal to value
      * "notEqual" - Not equal to value
      * "greaterThan" - Greater than value
      * "lessThan" - Less than value
      * "greaterThanOrEqual" - Greater than or equal to value
      * "lessThanOrEqual" - Less than or equal to value
    * `value1` - First numeric value for comparison
    * `value2` - Second numeric value for comparison (only required for "between" and "notBetween")
    * `allow_blank` - Whether to allow blank/empty values (default: true)
    * `error_title` - Optional title for error message when validation fails
    * `error_message` - Optional description for error message when validation fails
    * `prompt_title` - Optional title for input prompt
    * `prompt_message` - Optional description for input prompt

  ## Examples

  ```elixir
  # Allow only numbers between 1 and 10
  DataValidation.add_number_validation(
    spreadsheet,
    "Sheet1",
    "A1:A10",
    "between",
    1.0,
    10.0,
    true,
    "Invalid Number",
    "Please enter a number between 1 and 10",
    "Number Input",
    "Enter a number between 1 and 10"
  )

  # Allow only numbers greater than 100
  DataValidation.add_number_validation(
    spreadsheet,
    "Sheet1",
    "B1:B10",
    "greaterThan",
    100.0,
    nil,
    true,
    nil,
    nil,
    nil,
    nil
  )
  ```
  """
  @spec add_number_validation(
          Spreadsheet.t(),
          String.t(),
          String.t(),
          String.t(),
          float(),
          float() | nil,
          boolean(),
          String.t() | nil,
          String.t() | nil,
          String.t() | nil,
          String.t() | nil
        ) :: :ok | {:error, atom()}
  def add_number_validation(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_range,
        operator,
        value1,
        value2 \\ nil,
        allow_blank \\ true,
        error_title \\ nil,
        error_message \\ nil,
        prompt_title \\ nil,
        prompt_message \\ nil
      ) do
    # Convert numerical values to strings as required by the Rust NIF
    value1_str = Float.to_string(value1)
    value2_str = if value2, do: Float.to_string(value2), else: nil

    case UmyaNative.add_number_validation(
           UmyaSpreadsheet.unwrap_ref(ref),
           sheet_name,
           cell_range,
           operator,
           value1_str,
           value2_str,
           allow_blank,
           error_title,
           error_message,
           prompt_title,
           prompt_message
         ) do
      :ok -> :ok
      {:ok, :ok} -> :ok
      {:error, reason} -> {:error, reason}
    end
  end

  @doc """
  Adds date validation to a range of cells.

  This restricts input to dates that meet specified conditions.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet
    * `cell_range` - Range of cells in A1 notation (e.g., "A1:B10")
    * `operator` - Comparison operator, one of:
      * "between" - Between two dates
      * "notBetween" - Not between two dates
      * "equal" - Equal to date
      * "notEqual" - Not equal to date
      * "greaterThan" - After date
      * "lessThan" - Before date
      * "greaterThanOrEqual" - On or after date
      * "lessThanOrEqual" - On or before date
    * `date1` - First date for comparison in ISO format (YYYY-MM-DD)
    * `date2` - Second date for comparison (only required for "between" and "notBetween")
    * `allow_blank` - Whether to allow blank/empty values (default: true)
    * `error_title` - Optional title for error message when validation fails
    * `error_message` - Optional description for error message when validation fails
    * `prompt_title` - Optional title for input prompt
    * `prompt_message` - Optional description for input prompt

  ## Examples

  ```elixir
  # Allow only dates in 2023
  DataValidation.add_date_validation(
    spreadsheet,
    "Sheet1",
    "A1:A10",
    "between",
    "2023-01-01",
    "2023-12-31",
    true,
    "Invalid Date",
    "Please enter a date in 2023",
    "Date Input",
    "Enter a date in 2023"
  )

  # Allow only dates after today
  DataValidation.add_date_validation(
    spreadsheet,
    "Sheet1",
    "B1:B10",
    "greaterThan",
    Date.utc_today() |> Date.to_iso8601(),
    nil,
    true,
    "Invalid Date",
    "Please enter a future date",
    nil,
    nil
  )
  ```
  """
  @spec add_date_validation(
          Spreadsheet.t(),
          String.t(),
          String.t(),
          String.t(),
          String.t(),
          String.t() | nil,
          boolean(),
          String.t() | nil,
          String.t() | nil,
          String.t() | nil,
          String.t() | nil
        ) :: :ok | {:error, atom()}
  def add_date_validation(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_range,
        operator,
        date1,
        date2 \\ nil,
        allow_blank \\ true,
        error_title \\ nil,
        error_message \\ nil,
        prompt_title \\ nil,
        prompt_message \\ nil
      ) do
    case UmyaNative.add_date_validation(
           UmyaSpreadsheet.unwrap_ref(ref),
           sheet_name,
           cell_range,
           operator,
           date1,
           date2,
           allow_blank,
           error_title,
           error_message,
           prompt_title,
           prompt_message
         ) do
      :ok -> :ok
      {:ok, :ok} -> :ok
      {:error, reason} -> {:error, reason}
    end
  end

  @doc """
  Adds text length validation to a range of cells.

  This restricts input to text with a length that meets specified conditions.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet
    * `cell_range` - Range of cells in A1 notation (e.g., "A1:B10")
    * `operator` - Comparison operator, one of:
      * "between" - Between two lengths
      * "notBetween" - Not between two lengths
      * "equal" - Equal to length
      * "notEqual" - Not equal to length
      * "greaterThan" - Longer than length
      * "lessThan" - Shorter than length
      * "greaterThanOrEqual" - At least length characters
      * "lessThanOrEqual" - At most length characters
    * `length1` - First length value for comparison
    * `length2` - Second length value for comparison (only required for "between" and "notBetween")
    * `allow_blank` - Whether to allow blank/empty values (default: true)
    * `error_title` - Optional title for error message when validation fails
    * `error_message` - Optional description for error message when validation fails
    * `prompt_title` - Optional title for input prompt
    * `prompt_message` - Optional description for input prompt

  ## Examples

  ```elixir
  # Limit text to between 5 and 10 characters
  DataValidation.add_text_length_validation(
    spreadsheet,
    "Sheet1",
    "A1:A10",
    "between",
    5,
    10,
    true,
    "Invalid Text Length",
    "Please enter text between 5 and 10 characters",
    "Text Input",
    "Enter text between 5 and 10 characters"
  )

  # Limit text to at most 20 characters
  DataValidation.add_text_length_validation(
    spreadsheet,
    "Sheet1",
    "B1:B10",
    "lessThanOrEqual",
    20,
    nil,
    true,
    nil,
    nil,
    nil,
    nil
  )
  ```
  """
  @spec add_text_length_validation(
          Spreadsheet.t(),
          String.t(),
          String.t(),
          String.t(),
          integer(),
          integer() | nil,
          boolean(),
          String.t() | nil,
          String.t() | nil,
          String.t() | nil,
          String.t() | nil
        ) :: :ok | {:error, atom()}
  def add_text_length_validation(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_range,
        operator,
        length1,
        length2 \\ nil,
        allow_blank \\ true,
        error_title \\ nil,
        error_message \\ nil,
        prompt_title \\ nil,
        prompt_message \\ nil
      ) do
    case UmyaNative.add_text_length_validation(
           UmyaSpreadsheet.unwrap_ref(ref),
           sheet_name,
           cell_range,
           operator,
           length1,
           length2,
           allow_blank,
           error_title,
           error_message,
           prompt_title,
           prompt_message
         ) do
      :ok -> :ok
      {:ok, :ok} -> :ok
      {:error, reason} -> {:error, reason}
    end
  end

  @doc """
  Adds custom formula validation to a range of cells.

  This allows specifying a custom Excel formula for validation.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet
    * `cell_range` - Range of cells in A1 notation (e.g., "A1:B10")
    * `formula` - Excel formula that should evaluate to TRUE/FALSE (without the = sign)
    * `allow_blank` - Whether to allow blank/empty values (default: true)
    * `error_title` - Optional title for error message when validation fails
    * `error_message` - Optional description for error message when validation fails
    * `prompt_title` - Optional title for input prompt
    * `prompt_message` - Optional description for input prompt

  ## Examples

  ```elixir
  # Custom validation to ensure value is divisible by 3
  DataValidation.add_custom_validation(
    spreadsheet,
    "Sheet1",
    "A1:A10",
    "MOD(A1,3)=0",
    true,
    "Invalid Value",
    "Please enter a value divisible by 3",
    "Value Input",
    "Enter a value divisible by 3"
  )
  ```
  """
  @spec add_custom_validation(
          Spreadsheet.t(),
          String.t(),
          String.t(),
          String.t(),
          boolean(),
          String.t() | nil,
          String.t() | nil,
          String.t() | nil,
          String.t() | nil
        ) :: :ok | {:error, atom()}
  def add_custom_validation(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_range,
        formula,
        allow_blank \\ true,
        error_title \\ nil,
        error_message \\ nil,
        prompt_title \\ nil,
        prompt_message \\ nil
      ) do
    case UmyaNative.add_custom_validation(
           UmyaSpreadsheet.unwrap_ref(ref),
           sheet_name,
           cell_range,
           formula,
           allow_blank,
           error_title,
           error_message,
           prompt_title,
           prompt_message
         ) do
      :ok -> :ok
      {:ok, :ok} -> :ok
      {:error, reason} -> {:error, reason}
    end
  end

  @doc """
  Removes all data validation from a range of cells.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet
    * `cell_range` - Range of cells in A1 notation (e.g., "A1:B10")

  ## Examples

  ```elixir
  # Remove all validation from a range
  DataValidation.remove_data_validation(
    spreadsheet,
    "Sheet1",
    "A1:B10"
  )
  ```
  """
  @spec remove_data_validation(
          Spreadsheet.t(),
          String.t(),
          String.t()
        ) :: :ok | {:error, atom()}
  def remove_data_validation(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_range
      ) do
    case UmyaNative.remove_data_validation(
           UmyaSpreadsheet.unwrap_ref(ref),
           sheet_name,
           cell_range
         ) do
      :ok -> :ok
      {:ok, :ok} -> :ok
      {:error, reason} -> {:error, reason}
    end
  end
end
