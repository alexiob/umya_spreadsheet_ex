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
    * `date1` - First date for comparison as an ISO string (YYYY-MM-DD) or Elixir Date struct
    * `date2` - Second date for comparison as an ISO string or Date struct (only required for "between" and "notBetween")
    * `allow_blank` - Whether to allow blank/empty values (default: true)
    * `error_title` - Optional title for error message when validation fails
    * `error_message` - Optional description for error message when validation fails
    * `prompt_title` - Optional title for input prompt
    * `prompt_message` - Optional description for input prompt

  ## Examples

  ```elixir
  # Allow only dates in 2023 (using string dates)
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

  # Allow only dates in 2023 (using Date structs)
  DataValidation.add_date_validation(
    spreadsheet,
    "Sheet1",
    "A1:A10",
    "between",
    ~D[2023-01-01],
    ~D[2023-12-31],
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
    Date.utc_today(),  # Date struct works directly
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
          String.t() | Date.t(),
          (String.t() | Date.t()) | nil,
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
    # Convert Date structs to ISO strings if needed
    date1_str =
      case date1 do
        %Date{} -> Date.to_iso8601(date1)
        str when is_binary(str) -> str
      end

    date2_str =
      case date2 do
        %Date{} -> Date.to_iso8601(date2)
        str when is_binary(str) -> str
        nil -> nil
      end

    case UmyaNative.add_date_validation(
           UmyaSpreadsheet.unwrap_ref(ref),
           sheet_name,
           cell_range,
           operator,
           date1_str,
           date2_str,
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

  @doc """
  Gets all data validation rules for a sheet or range.

  Returns a list of maps, each containing details about a data validation rule.
  Each map includes the range it applies to and various other properties depending on the rule type.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet
    * `cell_range` - Optional range of cells to filter rules by (e.g., "A1:B10")

  ## Examples

  ```elixir
  # Get all validation rules in a sheet
  rules = DataValidation.get_data_validations(spreadsheet, "Sheet1")
  ```
  """
  @spec get_data_validations(
          Spreadsheet.t(),
          String.t(),
          String.t() | nil
        ) :: [map()]
  def get_data_validations(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_range \\ nil
      ) do
    UmyaNative.get_data_validations(
      UmyaSpreadsheet.unwrap_ref(ref),
      sheet_name,
      cell_range
    )
  end

  @doc """
  Gets list validation rules for a sheet or range.

  Returns list validation rules that restrict input to specified options.
  Each map includes the range it applies to, the list items, and various other properties.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet
    * `cell_range` - Optional range of cells to filter rules by (e.g., "A1:B10")

  ## Examples

  ```elixir
  # Get list validation rules in a sheet
  list_rules = DataValidation.get_list_validations(spreadsheet, "Sheet1")
  ```
  """
  @spec get_list_validations(
          Spreadsheet.t(),
          String.t(),
          String.t() | nil
        ) :: [map()]
  def get_list_validations(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_range \\ nil
      ) do
    UmyaNative.get_list_validations(
      UmyaSpreadsheet.unwrap_ref(ref),
      sheet_name,
      cell_range
    )
  end

  @doc """
  Gets number validation rules for a sheet or range.

  Returns number validation rules that restrict input to numeric values.
  Each map includes the range it applies to, operator, values, and various other properties.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet
    * `cell_range` - Optional range of cells to filter rules by (e.g., "A1:B10")

  ## Examples

  ```elixir
  # Get number validation rules in a sheet
  number_rules = DataValidation.get_number_validations(spreadsheet, "Sheet1")
  ```
  """
  @spec get_number_validations(
          Spreadsheet.t(),
          String.t(),
          String.t() | nil
        ) :: [map()]
  def get_number_validations(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_range \\ nil
      ) do
    UmyaNative.get_number_validations(
      UmyaSpreadsheet.unwrap_ref(ref),
      sheet_name,
      cell_range
    )
  end

  @doc """
  Gets date validation rules for a sheet or range.

  Returns date validation rules that restrict input to date values.
  Each map includes the range it applies to, operator, date values, and various other properties.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet
    * `cell_range` - Optional range of cells to filter rules by (e.g., "A1:B10")

  ## Examples

  ```elixir
  # Get date validation rules in a sheet
  date_rules = DataValidation.get_date_validations(spreadsheet, "Sheet1")
  ```
  """
  @spec get_date_validations(
          Spreadsheet.t(),
          String.t(),
          String.t() | nil
        ) :: [map()]
  def get_date_validations(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_range \\ nil
      ) do
    UmyaNative.get_date_validations(
      UmyaSpreadsheet.unwrap_ref(ref),
      sheet_name,
      cell_range
    )
  end

  @doc """
  Gets text length validation rules for a sheet or range.

  Returns text length validation rules that restrict input based on text length.
  Each map includes the range it applies to, operator, length values, and various other properties.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet
    * `cell_range` - Optional range of cells to filter rules by (e.g., "A1:B10")

  ## Examples

  ```elixir
  # Get text length validation rules in a sheet
  text_length_rules = DataValidation.get_text_length_validations(spreadsheet, "Sheet1")
  ```
  """
  @spec get_text_length_validations(
          Spreadsheet.t(),
          String.t(),
          String.t() | nil
        ) :: [map()]
  def get_text_length_validations(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_range \\ nil
      ) do
    UmyaNative.get_text_length_validations(
      UmyaSpreadsheet.unwrap_ref(ref),
      sheet_name,
      cell_range
    )
  end

  @doc """
  Gets custom formula validation rules for a sheet or range.

  Returns custom formula validation rules that restrict input based on a custom formula.
  Each map includes the range it applies to, formula, and various other properties.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet
    * `cell_range` - Optional range of cells to filter rules by (e.g., "A1:B10")

  ## Examples

  ```elixir
  # Get custom formula validation rules in a sheet
  custom_rules = DataValidation.get_custom_validations(spreadsheet, "Sheet1")
  ```
  """
  @spec get_custom_validations(
          Spreadsheet.t(),
          String.t(),
          String.t() | nil
        ) :: [map()]
  def get_custom_validations(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_range \\ nil
      ) do
    UmyaNative.get_custom_validations(
      UmyaSpreadsheet.unwrap_ref(ref),
      sheet_name,
      cell_range
    )
  end

  @doc """
  Checks if a sheet has any data validation rules.

  Returns `true` if the sheet has any validation rules, `false` otherwise.
  If a cell range is provided, only checks for rules applying to that range.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet
    * `cell_range` - Optional range of cells to check (e.g., "A1:B10")

  ## Examples

  ```elixir
  # Check if a sheet has any validation rules
  has_rules = DataValidation.has_data_validations(spreadsheet, "Sheet1")

  # Check if a specific range has validation rules
  has_rules_in_range = DataValidation.has_data_validations(spreadsheet, "Sheet1", "A1:A10")
  ```
  """
  @spec has_data_validations(
          Spreadsheet.t(),
          String.t(),
          String.t() | nil
        ) :: boolean()
  def has_data_validations(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_range \\ nil
      ) do
    UmyaNative.has_data_validations(
      UmyaSpreadsheet.unwrap_ref(ref),
      sheet_name,
      cell_range
    )
  end

  @doc """
  Counts the number of data validation rules in a sheet.

  Returns the count of validation rules in the sheet.
  If a cell range is provided, only counts rules applying to that range.

  ## Parameters

    * `spreadsheet` - A spreadsheet struct
    * `sheet_name` - Name of the sheet
    * `cell_range` - Optional range of cells to count rules for (e.g., "A1:B10")

  ## Examples

  ```elixir
  # Count all validation rules in a sheet
  rule_count = DataValidation.count_data_validations(spreadsheet, "Sheet1")

  # Count validation rules in a specific range
  range_rule_count = DataValidation.count_data_validations(spreadsheet, "Sheet1", "A1:A10")
  ```
  """
  @spec count_data_validations(
          Spreadsheet.t(),
          String.t(),
          String.t() | nil
        ) :: non_neg_integer()
  def count_data_validations(
        %Spreadsheet{reference: ref},
        sheet_name,
        cell_range \\ nil
      ) do
    UmyaNative.count_data_validations(
      UmyaSpreadsheet.unwrap_ref(ref),
      sheet_name,
      cell_range
    )
  end
end
