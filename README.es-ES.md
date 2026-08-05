

# UmyaSpreadsheet

[![CI](https://github.com/alexiob/umya_spreadsheet_ex/actions/workflows/ci.yml/badge.svg)](https://github.com/alexiob/umya_spreadsheet_ex/actions/workflows/ci.yml)
[![Hex.pm](https://img.shields.io/hexpm/v/umya_spreadsheet_ex.svg)](https://hex.pm/packages/umya_spreadsheet_ex)
[![Hex Docs](https://img.shields.io/badge/hex-docs-blue.svg)](https://hexdocs.pm/umya_spreadsheet_ex/)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

![Logotipo de UmyaSpreadsheetEx](assets/logo.png)

Todo lo que necesitabas (y probablemente más) para trabajar con archivos de Excel en Elixir, ¡impulsado por Rust!

Un envoltorio NIF de Elixir para la increíble biblioteca Rust [umya-spreadsheet](https://github.com/MathNya/umya-spreadsheet), que proporciona capacidades integrales de manipulación de archivos Excel (.xlsx, .xlsm) con los beneficios de rendimiento de Rust.

## Características

- **Crear y Manipular Hojas de Cálculo**

  - Crear nuevas hojas de cálculo desde cero
  - Leer y escribir archivos Excel existentes
  - Soporte para los formatos de archivo .xlsx y .xlsm
  - Funcionalidad de exportación a CSV
  - Opciones de escritor ligero para un mejor uso de la memoria

- **Operaciones con Celdas**

  - Obtener y establecer valores de celdas
  - Obtener valores con formato (números, monedas, fechas, etc.)
  - Mover rangos de celdas
  - Fusionar celdas
  - Eliminar celdas
  - Agregar y gestionar comentarios con información del autor

- **Estilos y Formato**

  - Aplicar colores (fondo, fuente)
  - Establecer propiedades de fuente (tamaño, negrita, cursiva, subrayado, tachado, nombre)
  - Tipografía avanzada con soporte para familia de fuentes y esquema de fuentes
  - Control de fuentes consciente del tema para un estilo de documento consistente
  - Formato de texto avanzado (varios estilos de subrayado, rotación de texto, sangría)
  - Estilos de bordes mejorados (segmentado, punteado, doble, etc.)
  - Aplicar formatos numéricos (moneda, porcentaje, fechas, etc.)
  - Habilitar ajuste de texto
  - Definir anchos de columna y altos de fila
  - Soporte para texto enriquecido con formato mixto dentro de las celdas
  - Crear y manipular elementos de texto con formato
  - Convertir texto enriquecido a/desde HTML
  - Formato condicional (reglas de valor de celda, escalas de color, barras de datos, reglas superior/inferior, reglas de texto)
  - Validación de datos (listas desplegadas, rangos numéricos, restricciones de fecha, límites de longitud de texto, fórmulas personalizadas)

- **Gestión de Hojas**

  - Agregar, clonar y eliminar hojas
  - Ocultar/mostrar hojas
  - Establecer visibilidad de cuadrícula
  - Mover y reordenar hojas
  - Configurar ajustes de ventana y pestaña activa
  - Establecer selección de celdas y vista de hoja de trabajo

- **Seguridad**

  - Protección con contraseña para libros de trabajo
  - Protección a nivel de hoja de cálculo
  - Opciones de escritor ligero para protección con contraseña

- **Operaciones de Filas y Columnas**

  - Insertar y eliminar filas y columnas
  - Ajustar altos de fila y anchos de columna
  - Aplicar estilo a filas/columnas enteras

- **Configuración de Impresión y Página**

  - Controlar orientación de página (vertical/horizontal)
  - Configurar tamaño de papel y escala
  - Establecer márgenes de página y encabezados/pies de página
  - Definir áreas de impresión y títulos (filas/columnas repetitivas)
  - Controlar centrado de impresión y opciones de ajuste a página

- **Soporte de Fórmulas**
  - Establecer fórmulas regulares en celdas individuales
  - Crear fórmulas matriciales en múltiples celdas
  - Definir rangos con nombre para referencias de fórmulas más fáciles
  - Crear nombres definidos para constantes y fórmulas
  - Listar nombres definidos en libros de trabajo

- **Organización y Análisis de Datos**
  - Agregar filtros automáticos a los encabezados de columna para filtrado interactivo
  - Habilitar/deshabilitar y gestionar filtros automáticos en hojas de cálculo
  - Consultar rangos y estados de filtros automáticos
  - Crear hojas de cálculo con vistas filtradas para un análisis de datos más fácil

- **Elementos Visuales**
  - Agregar imágenes (PNG, JPEG)
  - Crear gráficos (Línea, Barras, Circular y más)
  - Agregar formas (rectángulos, círculos, flechas, etc.)
  - Crear cuadros de texto para anotaciones
  - Conectar celdas con líneas conectoras
  - Posicionar elementos visuales con precisión

- **Operaciones Concurrentes**
  - Creación segura en hilos de múltiples hojas de cálculo
  - Soporte para operaciones de lectura/escritura en paralelo
  - Directrices para la manipulación concurrente de hojas de cálculo
  - Ejemplos de patrones seguros en hilos

- **Hipervínculos**
  - Agregar URL web, direcciones de correo electrónico, rutas de archivo y referencias internas
  - Obtener, actualizar y eliminar hipervínculos de celdas
  - Operaciones masivas de hipervínculos para una gestión eficiente
  - Integración con valores de celdas y sugerencias personalizadas

- **Análisis de Datos**
  - Crear y gestionar tablas dinámicas
  - Configurar campos de fila, columna y datos
  - Actualizar datos de tablas dinámicas
  - Posicionar tablas dinámicas en hojas de cálculo

- **Tablas de Excel**
  - Crear tablas estructuradas con encabezados y rangos de datos
  - Aplicar estilos de tabla integrados con opciones de personalización
  - Agregar, modificar y eliminar columnas de tabla dinámicamente
  - Configurar filas de totales con diversas funciones de cálculo
  - Gestionar capacidades de filtrado y ordenación de tablas
  - Soporte para metadatos de tabla y propiedades de columna

## Información de Versiones

Este paquete se basa en:

- **umya-spreadsheet**: v2.3.0 - Biblioteca robusta de Excel en Rust
- **Rustler**: v0.36.1 - Interoperabilidad Rust/Elixir sin problemas con excelente rendimiento

## Documentación

UmyaSpreadsheet tiene guías integrales para todas las funciones principales:

### Documentación en Línea

- [**Índice de Guías**](https://hexdocs.pm/umya_spreadsheet_ex/guides.html) - Punto de partida para toda la documentación
- [**Funciones de Fórmulas**](https://hexdocs.pm/umya_spreadsheet_ex/formula_functions.html) - Trabajo con fórmulas y referencias con nombre
- [**Filtros Automáticos**](https://hexdocs.pm/umya_spreadsheet_ex/auto_filters.html) - Creación y gestión de controles de filtro de Excel
- [**Configuración de Ventanas**](https://hexdocs.pm/umya_spreadsheet_ex/window_settings.html) - Controla cómo Excel muestra tus libros de trabajo
- [**Comentarios**](https://hexdocs.pm/umya_spreadsheet_ex/comments.html) - Agregar y gestionar comentarios de celdas
- [**Gráficos**](https://hexdocs.pm/umya_spreadsheet_ex/charts.html) - Creación y personalización de gráficos
- [**Validación de Datos**](https://hexdocs.pm/umya_spreadsheet_ex/data_validation.html) - Establecer reglas de entrada para celdas
- [**Opciones de Formato de Archivo**](https://hexdocs.pm/umya_spreadsheet_ex/file_format_options.html) - Control de compresión, cifrado y generación de formato binario

### Documentación de Referencia

- [**Guía de Solución de Problemas**](docs/troubleshooting.md) - Soluciones para problemas y errores comunes
- [**Limitaciones y Compatibilidad**](docs/limitations.md) - Limitaciones conocidas, matriz de compatibilidad y soluciones alternativas

## Instalación

Puedes instalarlo usando [igniter](https://hexdocs.pm/igniter) para la experiencia más cómoda:

```sh
# instala igniter si aún no lo has hecho
mix archive.install hex igniter_new
# luego instala umya_spreadsheet_ex
mix igniter.install umya_spreadsheet_ex
```

O agrega `umya_spreadsheet_ex` a tu lista de dependencias en `mix.exs`:

```elixir
def deps do
  [
    {:umya_spreadsheet_ex, "~> 0.7.0"}
  ]
end
```

El paquete incluye archivos NIF precompilados para plataformas comunes, pero se compilará desde el código fuente si es necesario.

Para forzar la compilación del NIF, establece la variable de entorno `UMYA_SPREADSHEET_BUILD` en `true`:

```bash
export UMYA_SPREADSHEET_BUILD=true

mix clean;
mix compile;
```

## Guía de Inicio Rápido

```elixir
# Create a new spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Write data to cells
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Hello")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "World")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C1", "42")

# Format your data
:ok = UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "A1", "FF0000") # Red
:ok = UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1", true)
:ok = UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "C1", "#,##0.00")

# Save your spreadsheet
:ok = UmyaSpreadsheet.write(spreadsheet, "hello_world.xlsx")

# Read it back
{:ok, loaded_spreadsheet} = UmyaSpreadsheet.read("hello_world.xlsx")
{:ok, value} = UmyaSpreadsheet.get_cell_value(loaded_spreadsheet, "Sheet1", "A1")
# => {:ok, "Hello"}
```

## Documentación y Guías

Para ejemplos más detallados y documentación completa de la API, visita:
[https://hexdocs.pm/umya_spreadsheet_ex](https://hexdocs.pm/umya_spreadsheet_ex)

Proporcionamos guías detalladas para funciones específicas:

- [**Gráficos**](https://hexdocs.pm/umya_spreadsheet_ex/charts.html) - Creación y personalización de varios tipos de gráficos
- [**Exportación a CSV y Rendimiento**](https://hexdocs.pm/umya_spreadsheet_ex/csv_export_and_performance.html) - Exportación CSV y escritores optimizados
- [**Validación de Datos**](https://hexdocs.pm/umya_spreadsheet_ex/data_validation.html) - Control y validación de valores de entrada de celdas
- [**Tablas de Excel**](https://hexdocs.pm/umya_spreadsheet_ex/excel_tables.html) - Creación, estilización y gestión de tablas estructuradas de Excel
- [**Opciones de Formato de Archivo**](https://hexdocs.pm/umya_spreadsheet_ex/file_format_options.html) - Control de compresión, cifrado y generación de Excel binario
- [**Manejo de Imágenes**](https://hexdocs.pm/umya_spreadsheet_ex/image_handling.html) - Trabajo con imágenes en hojas de cálculo
- [**Tablas Dinámicas**](https://hexdocs.pm/umya_spreadsheet_ex/pivot_tables.html) - Creación y gestión de tablas dinámicas de análisis de datos
- [**Configuración de Impresión**](https://hexdocs.pm/umya_spreadsheet_ex/print_settings.html) - Configuración de página y opciones de impresión
- [**Operaciones con Hojas**](https://hexdocs.pm/umya_spreadsheet_ex/sheet_operations.html) - Gestión efectiva de hojas de trabajo
- [**Estilos y Formato**](https://hexdocs.pm/umya_spreadsheet_ex/styling_and_formatting.html) - Haz que tus hojas de cálculo se vean profesionales

Puedes encontrar todas las guías en nuestro [Índice de Guías](https://hexdocs.pm/umya_spreadsheet_ex/guides.html).

## Ejemplos Completos de Uso

### Estilos y Formato

```elixir
# Create a styled spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Apply various styling options
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Styled Title")
:ok = UmyaSpreadsheet.set_font_size(spreadsheet, "Sheet1", "A1", 16)
:ok = UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1", true)
:ok = UmyaSpreadsheet.set_background_color(spreadsheet, "Sheet1", "A1", "CCFFCC") # Light green

# Add numbers with specific formats
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B3", "1234.56")
:ok = UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "B3", "#,##0.00")

:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B4", "0.42")
:ok = UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "B4", "0.00%")

# Currency formatting
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B5", "9999.99")
:ok = UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "B5", "$#,##0.00")

# Date formatting
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B6", "44500") # Excel date serial
:ok = UmyaSpreadsheet.set_number_format(spreadsheet, "Sheet1", "B6", "yyyy-mm-dd")

# Save the spreadsheet
:ok = UmyaSpreadsheet.write(spreadsheet, "formatted_spreadsheet.xlsx")
```

### Formato de Texto Enriquecido

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Create a rich text object
rich_text = UmyaSpreadsheet.RichText.create()

# Add formatted text with different styles
:ok = UmyaSpreadsheet.RichText.add_formatted_text(rich_text, "Bold text", %{bold: true})
:ok = UmyaSpreadsheet.RichText.add_formatted_text(rich_text, " and ", %{})
:ok = UmyaSpreadsheet.RichText.add_formatted_text(rich_text, "italic text", %{italic: true})
:ok = UmyaSpreadsheet.RichText.add_formatted_text(rich_text, " with ", %{})
:ok = UmyaSpreadsheet.RichText.add_formatted_text(rich_text, "colored text", %{color: "#FF0000"})

# Set the rich text to a cell
:ok = UmyaSpreadsheet.RichText.set_cell_rich_text(spreadsheet, "Sheet1", "A1", rich_text)

# Alternative: Create rich text from HTML
html_rich_text = UmyaSpreadsheet.RichText.create_from_html("<b>Bold</b> and <i>italic</i> text")
:ok = UmyaSpreadsheet.RichText.set_cell_rich_text(spreadsheet, "Sheet1", "A2", html_rich_text)

# Create individual text elements for more control
element1 = UmyaSpreadsheet.RichText.create_text_element("Large text", %{size: 18, bold: true})
element2 = UmyaSpreadsheet.RichText.create_text_element(" and small text", %{size: 10})

# Add elements to rich text
rich_text2 = UmyaSpreadsheet.RichText.create()
:ok = UmyaSpreadsheet.RichText.add_text_element(rich_text2, element1)
:ok = UmyaSpreadsheet.RichText.add_text_element(rich_text2, element2)
:ok = UmyaSpreadsheet.RichText.set_cell_rich_text(spreadsheet, "Sheet1", "A3", rich_text2)

# Generate HTML from rich text
html_output = UmyaSpreadsheet.RichText.to_html(rich_text)
# => "<b>Bold text</b> and <i>italic text</i> with <span style=\"color:#FF0000\">colored text</span>"

# Get font properties from text elements
{:ok, props} = UmyaSpreadsheet.RichText.get_element_font_properties(element1)
# props[:bold] => "true"
# props[:size] => "18"

:ok = UmyaSpreadsheet.write(spreadsheet, "rich_text_example.xlsx")
```

### Operaciones con Hojas

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add a new sheet
:ok = UmyaSpreadsheet.add_sheet(spreadsheet, "Data")

# Clone a sheet
:ok = UmyaSpreadsheet.clone_sheet(spreadsheet, "Sheet1", "Sheet1 Copy")

# Get all sheet names
sheet_names = UmyaSpreadsheet.get_sheet_names(spreadsheet)
# => ["Sheet1", "Data", "Sheet1 Copy"]

# Hide a sheet
:ok = UmyaSpreadsheet.set_sheet_state(spreadsheet, "Data", "hidden")

# Remove a sheet
:ok = UmyaSpreadsheet.remove_sheet(spreadsheet, "Sheet1 Copy")

# Turn off grid lines
:ok = UmyaSpreadsheet.set_show_grid_lines(spreadsheet, "Sheet1", false)
```

### Trabajo con Gráficos

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Prepare some data for the chart
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Month")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Sales")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "January")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", "1200")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "February")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B3", "1500")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A4", "March")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B4", "2200")

# Define data series for the chart
data_series = ["Sheet1!$B$2:$B$4"]
categories = "Sheet1!$A$2:$A$4"
title = "Quarterly Sales"

# Add a column chart
:ok = UmyaSpreadsheet.add_chart(
  spreadsheet,
  "Sheet1",
  "ColumnChart",  # Other options: LineChart, PieChart, BarChart
  "D2",           # Top-left position of chart
  "J10",          # Bottom-right position of chart
  title,          # Chart title
  data_series,    # Data series
  categories      # Categories (optional)
)
```

### Exportar a CSV

```elixir
# Read an existing spreadsheet
{:ok, spreadsheet} = UmyaSpreadsheet.read("sales_report.xlsx")

# Export just one sheet to CSV
:ok = UmyaSpreadsheet.write_csv(spreadsheet, "Sheet1", "sheet1_data.csv")
```

### Validación de Datos

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add a dropdown list to cells
:ok = UmyaSpreadsheet.add_list_validation(
  spreadsheet,
  "Sheet1",
  "A1:A10",
  ["Option 1", "Option 2", "Option 3"],
  true,  # Allow blank values
  "Invalid Selection",  # Error title
  "Please select from the list",  # Error message
  "Selection Required",  # Prompt title
  "Choose an option from the dropdown"  # Prompt message
)

# Add number validation (between 1 and 100)
:ok = UmyaSpreadsheet.add_number_validation(
  spreadsheet,
  "Sheet1",
  "B1:B10",
  "between",  # Operator: between, greaterThan, lessThan, etc.
  1.0,  # Minimum value
  100.0,  # Maximum value
  true,  # Allow blank values
  "Invalid Number",  # Error title
  "Please enter a number between 1 and 100"  # Error message
)

# Add date validation (future dates only)
# Option 1: Using Date struct directly
:ok = UmyaSpreadsheet.add_date_validation(
  spreadsheet,
  "Sheet1",
  "C1:C10",
  "greaterThan",  # Operator
  Date.utc_today(),  # Compare date as Date struct
  nil,  # Second date (for between operator)
  true,  # Allow blank values
  "Invalid Date",  # Error title
  "Please enter a future date"  # Error message
)

# Option 2: Using ISO string format
today_string = Date.utc_today() |> Date.to_iso8601()
:ok = UmyaSpreadsheet.add_date_validation(
  spreadsheet,
  "Sheet1",
  "D1:D10",
  "greaterThan",
  today_string,  # Compare date as string
  nil,
  true,
  "Invalid Date",
  "Please enter a future date"
)

# Add text length validation (max 10 characters)
:ok = UmyaSpreadsheet.add_text_length_validation(
  spreadsheet,
  "Sheet1",
  "D1:D10",
  "lessThanOrEqual",  # Operator
  10,  # Character limit
  nil,  # Second value (for between operator)
  true  # Allow blank values
)

# Remove validation from a range
:ok = UmyaSpreadsheet.remove_data_validation(
  spreadsheet,
  "Sheet1",
  "A5:A10"
)
```

### Uso de Memoria Mejorada con Escritores Ligeros

```elixir
# Create a large spreadsheet with many sheets
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# [... add lots of data ...]

# Use the light writer for better memory efficiency
:ok = UmyaSpreadsheet.write_light(spreadsheet, "large_file.xlsx")

# Or with password protection
:ok = UmyaSpreadsheet.write_with_password_light(spreadsheet, "secure_large_file.xlsx", "password123")
```

### Trabajo con Tablas Dinámicas

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Create some sample data for the pivot table
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Region")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Product")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C1", "Sales")

:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "North")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", "Apples")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C2", "10000")

:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "North")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B3", "Oranges")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C3", "8000")

:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A4", "South")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B4", "Apples")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C4", "12000")

# Add a sheet for the pivot table
:ok = UmyaSpreadsheet.add_sheet(spreadsheet, "Pivot")

# Create a pivot table
:ok = UmyaSpreadsheet.add_pivot_table(
  spreadsheet,                      # Spreadsheet object
  "Pivot",                          # Destination sheet
  "Sales Analysis",                 # Pivot table name
  "Sheet1",                         # Source sheet
  "A1:C4",                          # Source data range
  "A3",                             # Pivot table top-left position
  [0],                              # Row fields (Region - column index 0)
  [1],                              # Column fields (Product - column index 1)
  [{2, "sum", "Total Sales"}]       # Data fields (Sum of Sales)
)

# Check if a sheet has pivot tables
has_pivots = UmyaSpreadsheet.has_pivot_tables?(spreadsheet, "Pivot")
# => true

# Count pivot tables on a sheet
count = UmyaSpreadsheet.count_pivot_tables(spreadsheet, "Pivot")
# => 1

# Refresh all pivot tables
:ok = UmyaSpreadsheet.refresh_all_pivot_tables(spreadsheet)

# Remove a pivot table
:ok = UmyaSpreadsheet.remove_pivot_table(spreadsheet, "Pivot", "Sales Analysis")
```

### Trabajo con Tablas de Excel

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add some sample data for the table
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Product")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B1", "Category")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C1", "Price")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D1", "Stock")

UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "Laptop")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B2", "Electronics")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C2", 999.99)
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D2", 50)

UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A3", "Mouse")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "B3", "Electronics")
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "C3", 29.99)
UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "D3", 100)

# Create a table with the data
{:ok, :ok} = UmyaSpreadsheet.add_table(
  spreadsheet,
  "Sheet1",
  "ProductTable",
  "Product Inventory",
  "A1",
  "D3",
  ["Product", "Category", "Price", "Stock"],
  true  # Show totals row
)

# Apply a table style
{:ok, :ok} = UmyaSpreadsheet.set_table_style(
  spreadsheet,
  "Sheet1",
  "ProductTable",
  "TableStyleMedium9",
  true,   # Show first column
  false,  # Show last column
  true,   # Show banded rows
  false   # Show banded columns
)

# Add a new column to the table
{:ok, :ok} = UmyaSpreadsheet.add_table_column(
  spreadsheet,
  "Sheet1",
  "ProductTable",
  "Total Value",
  "sum",
  "Grand Total"
)

# Check if sheet has tables
{:ok, true} = UmyaSpreadsheet.has_tables?(spreadsheet, "Sheet1")

# Get all tables from the sheet
{:ok, tables} = UmyaSpreadsheet.get_tables(spreadsheet, "Sheet1")
[table | _] = tables
# table["name"] => "ProductTable"
# table["display_name"] => "Product Inventory"

# Remove the table
{:ok, :ok} = UmyaSpreadsheet.remove_table(spreadsheet, "Sheet1", "ProductTable")
```

## Características Avanzadas

### Operaciones de Filas y Columnas

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Insert 2 new rows at index 3
:ok = UmyaSpreadsheet.insert_new_row(spreadsheet, "Sheet1", 3, 2)

# Insert 1 new column at column C
:ok = UmyaSpreadsheet.insert_new_column(spreadsheet, "Sheet1", "C", 1)

# Remove row at index 5
:ok = UmyaSpreadsheet.remove_row(spreadsheet, "Sheet1", 5, 1)

# Set column width
:ok = UmyaSpreadsheet.set_column_width(spreadsheet, "Sheet1", "A", 15.5)

# Auto-fit column width based on content
:ok = UmyaSpreadsheet.set_column_auto_width(spreadsheet, "Sheet1", "B", true)

# Set row height
:ok = UmyaSpreadsheet.set_row_height(spreadsheet, "Sheet1", 1, 30.0)

# Apply styling to entire row
:ok = UmyaSpreadsheet.set_row_style(spreadsheet, "Sheet1", 1, "EEEEEE", "000000")
```

### Seguridad y Protección

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add data to the spreadsheet
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Confidential Data")
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A2", "Employee Salaries")

# Protect the entire workbook
:ok = UmyaSpreadsheet.set_workbook_protection(spreadsheet, true)

# Protect a specific worksheet
:ok = UmyaSpreadsheet.set_sheet_protection(
  spreadsheet,
  "Sheet1",
  true,
  "This sheet is protected"  # Optional message
)

# Save with password protection
:ok = UmyaSpreadsheet.write_with_password(spreadsheet, "confidential.xlsx", "secret123")

# Apply password to an existing file
:ok = UmyaSpreadsheet.set_password("original.xlsx", "protected_copy.xlsx", "secret123")
```

### Configuración de Impresión y Página

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add some content to the spreadsheet
:ok = UmyaSpreadsheet.set_cell_value(spreadsheet, "Sheet1", "A1", "Report Title")
:ok = UmyaSpreadsheet.set_font_size(spreadsheet, "Sheet1", "A1", 16)
:ok = UmyaSpreadsheet.set_font_bold(spreadsheet, "Sheet1", "A1", true)

# Configure print settings
# Set to landscape orientation
:ok = UmyaSpreadsheet.set_page_orientation(spreadsheet, "Sheet1", "landscape")

# Set to A4 paper size
:ok = UmyaSpreadsheet.set_paper_size(spreadsheet, "Sheet1", 9)

# Set page margins (in inches)
:ok = UmyaSpreadsheet.set_page_margins(spreadsheet, "Sheet1", 1.0, 0.75, 1.0, 0.75)

# Set header and footer margins
:ok = UmyaSpreadsheet.set_header_footer_margins(spreadsheet, "Sheet1", 0.5, 0.5)

# Add a custom header
:ok = UmyaSpreadsheet.set_header(spreadsheet, "Sheet1", "&C&\"Arial,Bold\"Confidential Report")

# Add a footer with page numbers
:ok = UmyaSpreadsheet.set_footer(spreadsheet, "Sheet1", "&RPage &P of &N")

# Define a specific print area
:ok = UmyaSpreadsheet.set_print_area(spreadsheet, "Sheet1", "A1:H20")

# Set rows 1-2 to repeat at the top of each printed page
:ok = UmyaSpreadsheet.set_print_titles(spreadsheet, "Sheet1", "1:2", "")

# Center the printout horizontally on the page
:ok = UmyaSpreadsheet.set_print_centered(spreadsheet, "Sheet1", true, false)

# Save the spreadsheet
:ok = UmyaSpreadsheet.write(spreadsheet, "print_ready_report.xlsx")
```

### Trabajo con Imágenes

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add an image to cell A1
:ok = UmyaSpreadsheet.add_image(spreadsheet, "Sheet1", "A1", "path/to/logo.png")

# When reading a spreadsheet with images, you can download them
{:ok, spreadsheet} = UmyaSpreadsheet.read("report_with_images.xlsx")
:ok = UmyaSpreadsheet.download_image(spreadsheet, "Sheet1", "A1", "downloaded_image.png")

# Replace an existing image
:ok = UmyaSpreadsheet.change_image(spreadsheet, "Sheet1", "A1", "path/to/new_logo.png")
```

### Trabajo con Objetos de Dibujo

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()

# Add a shape
:ok = UmyaSpreadsheet.add_shape(spreadsheet, "Sheet1", "D3", "rectangle", 200, 100, "blue", "black", 1.0)

# Add a text box
:ok = UmyaSpreadsheet.add_text_box(spreadsheet, "Sheet1", "D5", "Important Note", 200, 100, "yellow", "black", "gray", 1.0)

# Add a connector between cells
:ok = UmyaSpreadsheet.add_connector(spreadsheet, "Sheet1", "A1", "D3", "green", 1.5)
```

## Pruebas y Desarrollo

La biblioteca incluye casos de prueba integrales. Para ejecutarlos:

```bash
mix test
```

Los archivos de prueba se crean en el directorio `test/result_files` y son ignorados automáticamente por git.

Para más detalles sobre el desarrollo, consulta el archivo [DEVELOPMENT.md](DEVELOPMENT.md).

## Opciones Avanzadas de Formato de Archivo

```elixir
{:ok, spreadsheet} = UmyaSpreadsheet.new()
# Add data to the spreadsheet...

# Control the compression level (0-9)
:ok = UmyaSpreadsheet.write_with_compression(spreadsheet, "optimized.xlsx", 8)

# Enhanced encryption with AES256
:ok = UmyaSpreadsheet.write_with_encryption_options(
  spreadsheet,
  "secure.xlsx",
  "myPassword",
  "AES256",        # Algorithm
  "customSaltValue", # Optional salt value
  100000           # Optional spin count
)

# Generate binary XLSX for web responses
xlsx_binary = UmyaSpreadsheet.to_binary_xlsx(spreadsheet)

# In a Phoenix controller:
conn
|> put_resp_content_type("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
|> put_resp_header("content-disposition", ~s[attachment; filename="report.xlsx"])
|> send_resp(200, xlsx_binary)
```

## Consideraciones de Rendimiento

- Usa `lazy_read/1` para hojas de cálculo grandes para cargar hojas solo cuando se acceda a ellas
- Usa `write_light/2` y `write_with_password_light/3` para una mejor eficiencia de memoria con archivos grandes
- Para hojas de cálculo con muchas hojas, considera usar `new_empty/0` y agregar solo las hojas que necesites

## Licencia

UmyaSpreadsheetEx está disponible bajo la Licencia MIT. Consulta el archivo LICENSE para más información.

## Agradecimientos

Esta biblioteca es un envoltorio de la excelente biblioteca Rust [umya-spreadsheet](https://github.com/MathNya/umya-spreadsheet).
