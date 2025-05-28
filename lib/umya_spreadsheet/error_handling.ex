defmodule UmyaSpreadsheet.ErrorHandling do
  @moduledoc """
  Helper functions for standardizing error handling across the UmyaSpreadsheet library.
  """

  @doc """
  Unwraps nested error tuples to ensure a standardized error format.
  This helps with fixing the nested error tuples like {:error, {:error, reason}} to {:error, reason}
  """
  def unwrap_error({:error, {:error, reason}}), do: {:error, reason}
  def unwrap_error({:error, reason}), do: {:error, reason}
  def unwrap_error(other), do: other

  @doc """
  Standardizes success return values, ensuring they follow the {:ok, value} pattern.
  """
  def standardize_ok_result({:ok, :ok}), do: :ok
  def standardize_ok_result({:ok, value}), do: {:ok, value}
  def standardize_ok_result(:ok), do: :ok
  def standardize_ok_result(value) when not is_tuple(value), do: {:ok, value}
  def standardize_ok_result(other), do: other

  @doc """
  Standardizes result format for NIF functions that might return mixed formats.
  This is the main function to use for most function return value processing.
  """
  def standardize_result({:error, {:not_found, _}}), do: {:error, :not_found}
  def standardize_result({:error, _} = error), do: unwrap_error(error)
  def standardize_result({:ok, _} = success), do: standardize_ok_result(success)
  def standardize_result(:ok), do: :ok
  # Handle empty tuple from Rust NIF functions
  def standardize_result({}), do: :ok
  def standardize_result(value) when not is_tuple(value), do: {:ok, value}
  def standardize_result(other), do: other
end
