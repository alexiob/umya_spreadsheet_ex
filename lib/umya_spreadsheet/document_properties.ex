defmodule UmyaSpreadsheet.DocumentProperties do
  @moduledoc """
  Functions for managing document properties and custom metadata in spreadsheets.

  This module provides comprehensive functionality for:
  - Custom document properties (metadata)
  - Core document properties (title, description, subject, etc.)
  - Property management utilities

  Document properties allow you to store additional metadata about your spreadsheets,
  such as author information, project details, custom business data, and more.
  """

  alias UmyaSpreadsheet.Spreadsheet
  alias UmyaSpreadsheet.ErrorHandling
  alias UmyaNative

  # ========================================
  # Custom Properties Functions
  # ========================================

  @doc """
  Gets a custom document property value by name.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `property_name` - The name of the custom property

  ## Returns

  - `{:ok, value}` - The property value (string, number, boolean, or date)
  - `{:error, :not_found}` - If the property doesn't exist
  - `{:error, reason}` - On other failures

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, "Project Alpha"} = UmyaSpreadsheet.DocumentProperties.get_custom_property(spreadsheet, "ProjectName")
      {:error, :not_found} = UmyaSpreadsheet.DocumentProperties.get_custom_property(spreadsheet, "NonExistent")
  """
  def get_custom_property(%Spreadsheet{reference: ref}, property_name)
      when is_binary(property_name) do
    result =
      UmyaNative.get_custom_property(ref, property_name)
      |> ErrorHandling.standardize_result()

    # Convert string values to proper types
    case result do
      {:ok, value} when is_binary(value) -> {:ok, convert_property_value(value)}
      other -> other
    end
  end

  # Helper function to convert string values to appropriate types
  defp convert_property_value(value) when is_binary(value) do
    cond do
      value == "true" ->
        true

      value == "false" ->
        false

      # Try to convert to integer
      Regex.match?(~r/^\d+$/, value) ->
        case Integer.parse(value) do
          {int_val, ""} -> int_val
          _ -> value
        end

      # Try to convert to float (including negative)
      Regex.match?(~r/^-?\d+\.\d+$/, value) ->
        case Float.parse(value) do
          {float_val, ""} -> float_val
          _ -> value
        end

      # Handle scientific notation
      Regex.match?(~r/^-?\d+\.?\d*e[+-]?\d+$/i, value) ->
        case Float.parse(value) do
          {float_val, ""} -> float_val
          _ -> value
        end

      # Otherwise keep as string
      true ->
        value
    end
  end

  @doc """
  Sets a custom document property with a string value.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `property_name` - The name of the custom property
  - `value` - The string value to set

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.DocumentProperties.set_custom_property_string(spreadsheet, "ProjectName", "Project Alpha")
  """
  def set_custom_property_string(%Spreadsheet{reference: ref}, property_name, value)
      when is_binary(property_name) and is_binary(value) do
    UmyaNative.set_custom_property_string(ref, property_name, value)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets a custom document property with a numeric value.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `property_name` - The name of the custom property
  - `value` - The numeric value to set

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.DocumentProperties.set_custom_property_number(spreadsheet, "Version", 1.5)
      :ok = UmyaSpreadsheet.DocumentProperties.set_custom_property_number(spreadsheet, "Count", 42)
  """
  def set_custom_property_number(%Spreadsheet{reference: ref}, property_name, value)
      when is_binary(property_name) and is_number(value) do
    # The native function only supports integers, so we need to convert floats to string properties
    if is_float(value) do
      set_custom_property_string(
        %Spreadsheet{reference: ref},
        property_name,
        Float.to_string(value)
      )
    else
      UmyaNative.set_custom_property_number(ref, property_name, trunc(value))
      |> ErrorHandling.standardize_result()
    end
  end

  @doc """
  Sets a custom document property with a boolean value.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `property_name` - The name of the custom property
  - `value` - The boolean value to set

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.DocumentProperties.set_custom_property_bool(spreadsheet, "IsApproved", true)
      :ok = UmyaSpreadsheet.DocumentProperties.set_custom_property_bool(spreadsheet, "IsPublic", false)
  """
  def set_custom_property_bool(%Spreadsheet{reference: ref}, property_name, value)
      when is_binary(property_name) and is_boolean(value) do
    UmyaNative.set_custom_property_bool(ref, property_name, value)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets a custom document property with a date value.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `property_name` - The name of the custom property
  - `value` - The date value as a string (ISO 8601 format recommended)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.DocumentProperties.set_custom_property_date(spreadsheet, "ReleaseDate", "2024-12-31T00:00:00Z")
  """
  def set_custom_property_date(%Spreadsheet{reference: ref}, property_name, value)
      when is_binary(property_name) and is_binary(value) do
    # Parse the date string to extract year, month, and day
    with {:ok, datetime} <- parse_datetime(value) do
      year = datetime.year
      month = datetime.month
      day = datetime.day

      UmyaNative.set_custom_property_date(ref, property_name, year, month, day)
      |> ErrorHandling.standardize_result()
    else
      _ -> {:error, :invalid_date_format}
    end
  end

  @doc """
  Sets a custom document property with automatic type detection.

  This is a convenience function that automatically calls the appropriate
  type-specific function based on the value type.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `property_name` - The name of the custom property
  - `value` - The value to set (string, number, or boolean)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.DocumentProperties.set_custom_property(spreadsheet, "ProjectName", "Project Alpha")
      :ok = UmyaSpreadsheet.DocumentProperties.set_custom_property(spreadsheet, "Version", 1.5)
      :ok = UmyaSpreadsheet.DocumentProperties.set_custom_property(spreadsheet, "IsApproved", true)
  """
  def set_custom_property(spreadsheet, property_name, value) when is_binary(value) do
    set_custom_property_string(spreadsheet, property_name, value)
  end

  def set_custom_property(spreadsheet, property_name, value) when is_number(value) do
    if is_float(value) do
      set_custom_property_string(spreadsheet, property_name, Float.to_string(value))
    else
      set_custom_property_number(spreadsheet, property_name, value)
    end
  end

  def set_custom_property(spreadsheet, property_name, value) when is_boolean(value) do
    set_custom_property_bool(spreadsheet, property_name, value)
  end

  @doc """
  Removes a custom document property.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `property_name` - The name of the custom property to remove

  ## Returns

  - `:ok` on success (even if property didn't exist)
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.DocumentProperties.remove_custom_property(spreadsheet, "ProjectName")
  """
  def remove_custom_property(%Spreadsheet{reference: ref}, property_name)
      when is_binary(property_name) do
    UmyaNative.remove_custom_property(ref, property_name)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets a list of all custom property names.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct

  ## Returns

  - `{:ok, [property_names]}` - List of property names as strings
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, ["ProjectName", "Version", "IsApproved"]} = UmyaSpreadsheet.DocumentProperties.get_custom_property_names(spreadsheet)
  """
  def get_custom_property_names(%Spreadsheet{reference: ref}) do
    UmyaNative.get_custom_property_names(ref)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Checks if a custom property exists.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `property_name` - The name of the property to check

  ## Returns

  - `{:ok, true}` if the property exists
  - `{:ok, false}` if the property doesn't exist
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, true} = UmyaSpreadsheet.DocumentProperties.has_custom_property(spreadsheet, "ProjectName")
      {:ok, false} = UmyaSpreadsheet.DocumentProperties.has_custom_property(spreadsheet, "NonExistent")
  """
  def has_custom_property(%Spreadsheet{reference: ref}, property_name)
      when is_binary(property_name) do
    UmyaNative.has_custom_property(ref, property_name)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the count of custom properties.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct

  ## Returns

  - `{:ok, count}` - Number of custom properties
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, 3} = UmyaSpreadsheet.DocumentProperties.get_custom_properties_count(spreadsheet)
  """
  def get_custom_properties_count(%Spreadsheet{reference: ref}) do
    UmyaNative.get_custom_properties_count(ref)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Clears all custom properties.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.DocumentProperties.clear_custom_properties(spreadsheet)
  """
  def clear_custom_properties(%Spreadsheet{reference: ref}) do
    UmyaNative.clear_custom_properties(ref)
    |> ErrorHandling.standardize_result()
  end

  # ========================================
  # Core Document Properties Functions
  # ========================================

  @doc """
  Gets the document title.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct

  ## Returns

  - `{:ok, title}` - The document title as a string
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, "My Spreadsheet"} = UmyaSpreadsheet.DocumentProperties.get_title(spreadsheet)
  """
  def get_title(%Spreadsheet{reference: ref}) do
    UmyaNative.get_title(ref)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the document title.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `title` - The title to set

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.DocumentProperties.set_title(spreadsheet, "My Spreadsheet")
  """
  def set_title(%Spreadsheet{reference: ref}, title) when is_binary(title) do
    UmyaNative.set_title(ref, title)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the document description.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct

  ## Returns

  - `{:ok, description}` - The document description as a string
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, "Sales data for Q4"} = UmyaSpreadsheet.DocumentProperties.get_description(spreadsheet)
  """
  def get_description(%Spreadsheet{reference: ref}) do
    UmyaNative.get_description(ref)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the document description.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `description` - The description to set

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.DocumentProperties.set_description(spreadsheet, "Sales data for Q4")
  """
  def set_description(%Spreadsheet{reference: ref}, description) when is_binary(description) do
    UmyaNative.set_description(ref, description)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the document subject.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct

  ## Returns

  - `{:ok, subject}` - The document subject as a string
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, "Financial Analysis"} = UmyaSpreadsheet.DocumentProperties.get_subject(spreadsheet)
  """
  def get_subject(%Spreadsheet{reference: ref}) do
    UmyaNative.get_subject(ref)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the document subject.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `subject` - The subject to set

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.DocumentProperties.set_subject(spreadsheet, "Financial Analysis")
  """
  def set_subject(%Spreadsheet{reference: ref}, subject) when is_binary(subject) do
    UmyaNative.set_subject(ref, subject)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the document keywords.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct

  ## Returns

  - `{:ok, keywords}` - The document keywords as a string
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, "sales, data, analysis"} = UmyaSpreadsheet.DocumentProperties.get_keywords(spreadsheet)
  """
  def get_keywords(%Spreadsheet{reference: ref}) do
    UmyaNative.get_keywords(ref)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the document keywords.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `keywords` - The keywords to set (typically comma-separated)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.DocumentProperties.set_keywords(spreadsheet, "sales, data, analysis")
  """
  def set_keywords(%Spreadsheet{reference: ref}, keywords) when is_binary(keywords) do
    UmyaNative.set_keywords(ref, keywords)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the document creator.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct

  ## Returns

  - `{:ok, creator}` - The document creator as a string
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, "John Doe"} = UmyaSpreadsheet.DocumentProperties.get_creator(spreadsheet)
  """
  def get_creator(%Spreadsheet{reference: ref}) do
    UmyaNative.get_creator(ref)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the document creator.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `creator` - The creator name to set

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.DocumentProperties.set_creator(spreadsheet, "John Doe")
  """
  def set_creator(%Spreadsheet{reference: ref}, creator) when is_binary(creator) do
    UmyaNative.set_creator(ref, creator)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the document last modified by.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct

  ## Returns

  - `{:ok, last_modified_by}` - The last modified by as a string
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, "Jane Smith"} = UmyaSpreadsheet.DocumentProperties.get_last_modified_by(spreadsheet)
  """
  def get_last_modified_by(%Spreadsheet{reference: ref}) do
    UmyaNative.get_last_modified_by(ref)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the document last modified by.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `last_modified_by` - The last modified by name to set

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.DocumentProperties.set_last_modified_by(spreadsheet, "Jane Smith")
  """
  def set_last_modified_by(%Spreadsheet{reference: ref}, last_modified_by)
      when is_binary(last_modified_by) do
    UmyaNative.set_last_modified_by(ref, last_modified_by)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the document category.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct

  ## Returns

  - `{:ok, category}` - The document category as a string
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, "Business"} = UmyaSpreadsheet.DocumentProperties.get_category(spreadsheet)
  """
  def get_category(%Spreadsheet{reference: ref}) do
    UmyaNative.get_category(ref)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the document category.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `category` - The category to set

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.DocumentProperties.set_category(spreadsheet, "Business")
  """
  def set_category(%Spreadsheet{reference: ref}, category) when is_binary(category) do
    UmyaNative.set_category(ref, category)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the document company.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct

  ## Returns

  - `{:ok, company}` - The document company as a string
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, "Acme Corp"} = UmyaSpreadsheet.DocumentProperties.get_company(spreadsheet)
  """
  def get_company(%Spreadsheet{reference: ref}) do
    UmyaNative.get_company(ref)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the document company.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `company` - The company name to set

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.DocumentProperties.set_company(spreadsheet, "Acme Corp")
  """
  def set_company(%Spreadsheet{reference: ref}, company) when is_binary(company) do
    UmyaNative.set_company(ref, company)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the document manager.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct

  ## Returns

  - `{:ok, manager}` - The document manager as a string
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, "Bob Johnson"} = UmyaSpreadsheet.DocumentProperties.get_manager(spreadsheet)
  """
  def get_manager(%Spreadsheet{reference: ref}) do
    UmyaNative.get_manager(ref)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the document manager.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `manager` - The manager name to set

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.DocumentProperties.set_manager(spreadsheet, "Bob Johnson")
  """
  def set_manager(%Spreadsheet{reference: ref}, manager) when is_binary(manager) do
    UmyaNative.set_manager(ref, manager)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the document created date.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct

  ## Returns

  - `{:ok, created}` - The created date as a string
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, "2024-01-01T00:00:00Z"} = UmyaSpreadsheet.DocumentProperties.get_created(spreadsheet)
  """
  def get_created(%Spreadsheet{reference: ref}) do
    UmyaNative.get_created(ref)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the document created date.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `created` - The created date to set (ISO 8601 format recommended)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.DocumentProperties.set_created(spreadsheet, "2024-01-01T00:00:00Z")
  """
  def set_created(%Spreadsheet{reference: ref}, created) when is_binary(created) do
    UmyaNative.set_created(ref, created)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Gets the document modified date.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct

  ## Returns

  - `{:ok, modified}` - The modified date as a string
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, "2024-05-15T12:30:00Z"} = UmyaSpreadsheet.DocumentProperties.get_modified(spreadsheet)
  """
  def get_modified(%Spreadsheet{reference: ref}) do
    UmyaNative.get_modified(ref)
    |> ErrorHandling.standardize_result()
  end

  @doc """
  Sets the document modified date.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `modified` - The modified date to set (ISO 8601 format recommended)

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      :ok = UmyaSpreadsheet.DocumentProperties.set_modified(spreadsheet, "2024-05-15T12:30:00Z")
  """
  def set_modified(%Spreadsheet{reference: ref}, modified) when is_binary(modified) do
    UmyaNative.set_modified(ref, modified)
    |> ErrorHandling.standardize_result()
  end

  # ========================================
  # Convenience Functions
  # ========================================

  # Helper function to parse datetime strings
  defp parse_datetime(datetime_str) do
    # Try common ISO formats
    case DateTime.from_iso8601(datetime_str) do
      {:ok, datetime, _offset} ->
        {:ok, datetime}

      _ ->
        # Try simple date format YYYY-MM-DD
        case Date.from_iso8601(datetime_str) do
          {:ok, date} -> {:ok, DateTime.new!(date, ~T[00:00:00])}
          _ -> :error
        end
    end
  end

  @doc """
  Sets multiple document properties at once.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct
  - `properties` - A map of property names to values

  ## Returns

  - `:ok` on success
  - `{:error, reason}` on first failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      properties = %{
        title: "My Report",
        description: "Monthly sales report",
        creator: "John Doe",
        company: "Acme Corp"
      }
      :ok = UmyaSpreadsheet.DocumentProperties.set_properties(spreadsheet, properties)
  """
  def set_properties(%Spreadsheet{} = spreadsheet, properties) when is_map(properties) do
    Enum.reduce_while(properties, :ok, fn {key, value}, _acc ->
      result =
        case key do
          :title -> set_title(spreadsheet, value)
          :description -> set_description(spreadsheet, value)
          :subject -> set_subject(spreadsheet, value)
          :keywords -> set_keywords(spreadsheet, value)
          :creator -> set_creator(spreadsheet, value)
          :last_modified_by -> set_last_modified_by(spreadsheet, value)
          :category -> set_category(spreadsheet, value)
          :company -> set_company(spreadsheet, value)
          :manager -> set_manager(spreadsheet, value)
          :created -> set_created(spreadsheet, value)
          :modified -> set_modified(spreadsheet, value)
          _ -> {:error, {:unknown_property, key}}
        end

      case result do
        :ok -> {:cont, :ok}
        error -> {:halt, error}
      end
    end)
  end

  @doc """
  Gets all core document properties as a map.

  ## Parameters

  - `spreadsheet` - The spreadsheet struct

  ## Returns

  - `{:ok, properties_map}` - Map of all core properties
  - `{:error, reason}` on failure

  ## Examples

      {:ok, spreadsheet} = UmyaSpreadsheet.read_file("input.xlsx")
      {:ok, properties} = UmyaSpreadsheet.DocumentProperties.get_all_properties(spreadsheet)
      # properties = %{
      #   title: "My Report",
      #   description: "Monthly sales report",
      #   creator: "John Doe",
      #   ...
      # }
  """
  def get_all_properties(%Spreadsheet{} = spreadsheet) do
    with {:ok, title} <- get_title(spreadsheet),
         {:ok, description} <- get_description(spreadsheet),
         {:ok, subject} <- get_subject(spreadsheet),
         {:ok, keywords} <- get_keywords(spreadsheet),
         {:ok, creator} <- get_creator(spreadsheet),
         {:ok, last_modified_by} <- get_last_modified_by(spreadsheet),
         {:ok, category} <- get_category(spreadsheet),
         {:ok, company} <- get_company(spreadsheet),
         {:ok, manager} <- get_manager(spreadsheet),
         {:ok, created} <- get_created(spreadsheet),
         {:ok, modified} <- get_modified(spreadsheet) do
      {:ok,
       %{
         title: title,
         description: description,
         subject: subject,
         keywords: keywords,
         creator: creator,
         last_modified_by: last_modified_by,
         category: category,
         company: company,
         manager: manager,
         created: created,
         modified: modified
       }}
    end
  end
end
