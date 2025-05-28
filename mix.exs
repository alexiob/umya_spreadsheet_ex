defmodule UmyaSpreadsheet.MixProject do
  use Mix.Project

  def project do
    [
      app: :umya_spreadsheet_ex,
      version: "0.6.9",
      elixir: "~> 1.14",
      start_permanent: Mix.env() == :prod,
      deps: deps(),
      rustler_crates: rustler_crates(),
      description: description(),
      package: package(),
      docs: docs()
    ]
  end

  # Run "mix help compile.app" to learn about applications.
  def application do
    [
      extra_applications: [:logger]
    ]
  end

  # Run "mix help deps" to learn about dependencies.
  defp deps do
    [
      {:rustler_precompiled, "~> 0.4"},
      {:rustler, "~> 0.36.1", runtime: false, optional: true},
      {:ex_doc, "~> 0.29", only: :dev, runtime: false},
      {:igniter, "~> 0.6", optional: true}
    ]
  end

  def rustler_crates do
    [
      umya_native: [
        path: "native/umya_native",
        mode: if(Mix.env() == :prod, do: :release, else: :debug)
      ]
    ]
  end

  defp description do
    """
    Elixir NIF wrapper for the umya-spreadsheet Rust library, providing Excel (.xlsx) file
    manipulation capabilities.
    """
  end

  defp package do
    [
      files: ~w(lib native .formatter.exs mix.exs README.md LICENSE),
      licenses: ["MIT"],
      links: %{
        "GitHub" => "https://github.com/alexiob/umya_spreadsheet_ex"
      },
      maintainers: ["Alessandro Iob"],
      organization: nil,
      submitter: "alexiob",
      maintainer: "alessandro@iob.dev"
    ]
  end

  defp docs do
    [
      main: "readme",
      extras: [
        "README.md",
        "CHANGELOG.md",
        "DEVELOPMENT.md",
        "LICENSE",
        "docs/guides.md",
        "docs/advanced_cell_formatting.md",
        "docs/advanced_sheet_features.md",
        "docs/auto_filters.md",
        "docs/charts.md",
        "docs/comments.md",
        "docs/conditional_formatting.md",
        "docs/csv_export_and_performance.md",
        "docs/data_validation.md",
        "docs/file_format_options.md",
        "docs/formula_functions.md",
        "docs/hyperlinks.md",
        "docs/image_handling.md",
        "docs/ole_objects.md",
        "docs/page_breaks.md",
        "docs/pivot_tables.md",
        "docs/print_settings.md",
        "docs/shapes_and_drawing.md",
        "docs/sheet_operations.md",
        "docs/styling_and_formatting.md",
        "docs/thread_safety.md",
        "docs/troubleshooting.md",
        "docs/limitations.md",
        "docs/window_settings.md"
      ],
      groups_for_extras: [
        Guides: [
          "docs/guides.md"
        ],
        "Basic Features": [
          "docs/sheet_operations.md",
          "docs/styling_and_formatting.md",
          "docs/csv_export_and_performance.md",
          "docs/file_format_options.md"
        ],
        "Advanced Features": [
          "docs/advanced_cell_formatting.md",
          "docs/advanced_sheet_features.md",
          "docs/conditional_formatting.md",
          "docs/data_validation.md",
          "docs/auto_filters.md",
          "docs/formula_functions.md",
          "docs/hyperlinks.md",
          "docs/ole_objects.md"
        ],
        "Visual Elements": [
          "docs/charts.md",
          "docs/image_handling.md",
          "docs/shapes_and_drawing.md",
          "docs/comments.md"
        ],
        "Layout & Printing": [
          "docs/page_breaks.md",
          "docs/pivot_tables.md",
          "docs/print_settings.md",
          "docs/window_settings.md"
        ],
        Technical: [
          "docs/thread_safety.md"
        ],
        Reference: [
          "docs/troubleshooting.md",
          "docs/limitations.md"
        ],
        Development: [
          "DEVELOPMENT.md",
          "CHANGELOG.md"
        ]
      ],
      source_url: "https://github.com/alexiob/umya_spreadsheet_ex",
      groups_for_modules: [
        API: [
          UmyaSpreadsheet
        ],
        Internal: [
          UmyaNative,
          UmyaSpreadsheet.Spreadsheet
        ]
      ]
    ]
  end
end
