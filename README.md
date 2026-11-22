# DataExport

[![NuGet](https://img.shields.io/nuget/v/DataExport)](https://www.nuget.org/packages/DataExport)
[![Build](https://github.com/Shahriyor187/DataExport/actions/workflows/ci.yml/badge.svg)](https://github.com/Shahriyor187/DataExport/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

A lightweight .NET library for exporting dictionary-based records to **JSON, XML, CSV, and Excel (XLSX)** files.

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

The stream remains open after export, so web applications can write straight to the response.

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
