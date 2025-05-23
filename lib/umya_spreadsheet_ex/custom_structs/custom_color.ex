defmodule UmyaSpreadsheetEx.CustomStructs.CustomColor do
  @moduledoc """
  Struct for representing colors with ARGB values.
  """
  
  defstruct [:argb]
  
  @doc """
  Creates a new CustomColor struct from a hex color string.
  
  ## Examples
  
      iex> UmyaSpreadsheetEx.CustomStructs.CustomColor.from_hex("#FF0000")
      %UmyaSpreadsheetEx.CustomStructs.CustomColor{argb: "FFFF0000"}
      
      iex> UmyaSpreadsheetEx.CustomStructs.CustomColor.from_hex("00FF00")
      %UmyaSpreadsheetEx.CustomStructs.CustomColor{argb: "FF00FF00"}
  """
  def from_hex(hex_color) when is_binary(hex_color) do
    hex = String.replace(hex_color, "#", "")
    argb = case String.length(hex) do
      6 -> "FF" <> String.upcase(hex)  # Add alpha channel if not present
      8 -> String.upcase(hex)          # Already has alpha channel
      _ -> "FFFFFFFF"                  # Default to white if invalid
    end
    %__MODULE__{argb: argb}
  end
  
  def from_hex(nil), do: nil
end
