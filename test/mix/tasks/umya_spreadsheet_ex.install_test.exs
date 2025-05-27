defmodule Mix.Tasks.UmyaSpreadsheetEx.InstallTest do
  use ExUnit.Case, async: true
  import Igniter.Test

  test "installation is idempotent" do
    assert {:ok, _igniter, %{warnings: [], notices: []}} =
             test_project()
             |> Igniter.compose_task("umya_spreadsheet_ex.install")
             |> apply_igniter!()
             |> Igniter.compose_task("umya_spreadsheet_ex.install")
             |> assert_unchanged()
             |> apply_igniter()
  end
end
