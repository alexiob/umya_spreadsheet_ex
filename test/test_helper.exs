# Clean up all test result files before starting the test suite
result_files_path = Path.join([File.cwd!(), "test", "result_files"])

if File.exists?(result_files_path) do
  result_files_path
  |> File.ls!()
  |> Enum.each(fn file ->
    file_path = Path.join(result_files_path, file)

    if File.regular?(file_path) do
      File.rm!(file_path)
    end
  end)
end

ExUnit.start()
