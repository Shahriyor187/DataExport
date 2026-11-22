# DataExport

[![NuGet](https://img.shields.io/nuget/v/DataExport)](https://www.nuget.org/packages/DataExport)
[![Build](https://github.com/Shahriyor187/DataExport/actions/workflows/ci.yml/badge.svg)](https://github.com/Shahriyor187/DataExport/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

A lightweight .NET library for exporting **C# objects, dictionaries, and DataTables** to **JSON, XML, CSV, and Excel (XLSX)**, with configurable output and streaming support.

## Low-memory Excel exports (experimental v2 foundation)

For large database queries, use the **forward-only** Open XML writer instead of `FromObjects(...)` or the existing ClosedXML workbook methods. The API consumes rows incrementally, preserves numeric/Boolean/date values and supports automatic sheet rollover. Supply `Columns` for empty sources or records with inconsistent properties. Memory use is bounded with respect to the number of rows, not necessarily the number or size of individual cells.

```csharp
await using var output = File.Create("students.xlsx");
await DataExporter.WriteExcelStreamingAsync(
    studentRows, // IAsyncEnumerable<Student>
    output,
    new StreamingExcelOptions
    {
        Columns = ["Name", "Score"],
        IncludeHeaders = true,
        SplitLargeWorksheets = true
    });
```

An overload accepts a live `DbDataReader`. The writer does not dispose your reader or destination stream. The streaming route does not offer ClosedXML-only features such as automatic column sizing, charts, or Excel formulas. Test large exports before using it in production.

## Multi-sheet Excel report builder

Use `ExcelReportBuilder` to create styled workbooks from several datasets. It supports nested property mapping, custom display headers and number formats, frozen headers, filters, alternating row styles and formula-based total rows.

```csharp
using var report = new ExcelReportBuilder();
report.AddSheet("Students", students, new ExcelReportSheetOptions
{
    Columns =
    [
        new ExcelReportColumn("Name", "Student"),
        new ExcelReportColumn("Score", "SAT score", "#,##0"),
        new ExcelReportColumn("Address.City", "City")
    ]
});
report.AddTotals("Students", 2);
report.AddSheet("Overview", new[] { new { Metric = "Total students", Value = 100 } });
report.SaveAs("report.xlsx");
```

This report builder uses an **in-memory workbook**, so use `WriteExcelStreamingAsync` instead for very large exports.

## Requirements

- .NET 8 or .NET 10

## Installation

```bash
dotnet add package DataExport
```

## Quick start

```csharp
using DataExportPackage;

var records = new List<Dictionary<string, object>>
{
    new()
    {
        ["Name"] = "Amina",
        ["Age"] = 21,
        ["Country"] = "Uzbekistan"
    },
    new()
    {
        ["Name"] = "Ali",
        ["Age"] = 24,
        ["Country"] = "Uzbekistan"
    }
};

var exporter = new DataExporter(records);

exporter.ToJson("people.json");
exporter.ToXml("people.xml");
exporter.ToCsv("people.csv");
exporter.ToExcel("people.xlsx");
```

## Supported formats

| Method | Output | Notes |
| --- | --- | --- |
| `ToJson(path)` | JSON | Indented JSON, serialized with `System.Text.Json`. |
| `ToXml(path)` | XML | Creates a `data` root with one `entry` element per record. Invalid XML key characters are encoded. |
| `ToCsv(path)` | CSV | UTF-8 without BOM, comma-delimited, with quoted fields escaped. |
| `ToExcel(path)` | XLSX | Creates a worksheet named `Data`; no header row by default, preserving the original API behavior. |

CSV and Excel use the union of all dictionary keys, in the order those keys are first encountered. CSV includes a header row by default. If a record does not contain a particular key, the corresponding cell is empty.

To include an Excel header row:

```csharp
exporter.ToExcel("people-with-headers.xlsx", includeHeaders: true);
```

## Notes

- The library writes to the file path you provide. The destination directory must already exist.
- Values in CSV, XML, and Excel are formatted using invariant culture where applicable.
- Null values are exported as empty strings in CSV, XML, and Excel. JSON preserves null values.
- The package does not currently export Word or PDF files.

## License

MIT. See [LICENSE](LICENSE).

## New in 1.2.0

### Export C# objects and DataTables

```csharp
var users = new[] { new { Name = "Amina", Score = 1450 } };
DataExporter.Export(users, "users.csv", ExportFormat.Csv);
var exporter = DataExporter.FromObjects(users);
var fromTable = DataExporter.FromDataTable(myDataTable);
```

### Export directly to streams

The stream remains open after export, so web applications can write straight to the response. The CSV encoding can be specified with `DataExportOptions.CsvEncoding` (UTF-8 without BOM by default).

```csharp
var exporter = DataExporter.FromObjects(users);
await exporter.ExportAsync(httpContext.Response.Body, ExportFormat.Json,
    cancellationToken: httpContext.RequestAborted);
```

For very large CSV exports, use `DataExporter.WriteCsvAsync(IAsyncEnumerable<T>, stream)`. It writes rows as they arrive without holding the entire dataset in memory. For sparse dictionaries, provide an explicit `Columns` list.

### Configure CSV and Excel output

```csharp
var options = new DataExportOptions
{
    Columns = ["Name", "Score"],
    CsvDelimiter = ';',
    CsvEncoding = System.Text.Encoding.UTF8,
    CsvProtectFormulas = true,
    WorksheetName = "Results",
    ExcelIncludeHeaders = true,
    ExcelFreezeHeader = true,
    ExcelAutoFilter = true,
    ExcelAutoFitColumns = false
};
exporter.ToCsv("users.csv", options);
exporter.ToExcel("users.xlsx", options);
```

Excel writes numeric, Boolean and date values using native cell types. The original `ToExcel(path)` still omits the header row for compatibility. The options-based CSV methods protect against formula-like values by default. The original `ToCsv(path)` preserves its previous unmodified output for compatibility. For untrusted data, prefer `ToCsv(path, new DataExportOptions())` or `ToCsvAsync`.

### ASP.NET Core download example

```csharp
app.MapGet("/export/users", async (HttpContext context) =>
{
    var users = new[] { new { Name = "Amina", Score = 1450 } };
    context.Response.ContentType = "text/csv; charset=utf-8";
    context.Response.Headers.ContentDisposition = "attachment; filename=users.csv";
    await DataExporter.FromObjects(users)
        .ExportAsync(context.Response.Body, ExportFormat.Csv,
            cancellationToken: context.RequestAborted);
});
```

This package doesn't provide an HTTP endpoint or require ASP.NET Core itself.
