# DataExport

[![NuGet](https://img.shields.io/nuget/v/DataExport)](https://www.nuget.org/packages/DataExport)
[![Build](https://github.com/Shahriyor187/DataExport/actions/workflows/ci.yml/badge.svg)](https://github.com/Shahriyor187/DataExport/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

A lightweight .NET library for exporting **C# objects, dictionaries, and DataTables** to **JSON, XML, CSV, and Excel (XLSX)**, with configurable output and streaming support.

> **Versioning:** Install **DataExport 2.0.0 or newer** to use the v2 features described below. GitHub source and the latest actually available NuGet package can differ, so check the installed package version. Upgrading from 1.x may require changes: `DataExporter` is now sealed and JSON serialization uses `System.Text.Json` rather than Newtonsoft.Json.

> **Runtime lifecycle:** The package supports `net8.0` for compatibility, but Microsoft ended support for the .NET 8 runtime on November 10, 2026. Prefer supported .NET 10 for new production applications.


## Low-memory Excel exports (v2)

For large database queries, use the **direct-ZIP, forward-only** Open XML writer instead of `FromObjects(...)` or the existing ClosedXML workbook methods. The API consumes rows incrementally, preserves numeric/Boolean/date values and supports automatic sheet rollover. Supply `Columns` for empty sources or records with inconsistent properties. Rows and worksheet XML are written incrementally into ZIP entries without an in-memory Open XML package graph. The output stream only needs to be writable, so write-only file streams and non-seekable HTTP response streams are supported. Benchmark the workload before making production memory guarantees.

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

An overload accepts a live `DbDataReader`. The writer does not dispose your reader or destination stream. On cancellation or failure the partly written XLSX must be discarded. The streaming route does not offer ClosedXML-only features such as automatic column sizing, charts, or Excel formulas. Test large exports before using it in production.

## Validate performance and large files

The repository includes a separate BenchmarkDotNet project and a forward-only
XLSX stress runner. The stress runner generates rows on demand, writes a real
workbook to a FileStream using a direct ZIP writer, then verifies **every worksheet's physical row count** using a
streaming XML reader—without opening the file in ClosedXML.

```bash
# Small reproduction / CI smoke test
dotnet run -c Release --project benchmarks/DataExport.Benchmarks -- --stress --rows 10000 --max-rows-per-sheet 500000

# 100k records
dotnet run -c Release --project benchmarks/DataExport.Benchmarks -- --stress --rows 100000 --results perf-100k.json

# 1M records, forced worksheet rollover; optional peak limit
dotnet run -c Release --project benchmarks/DataExport.Benchmarks -- --stress --rows 1000000 --max-rows-per-sheet 500000 --max-peak-mb 512 --results perf-1m.json

# Statistical benchmark of in-memory vs forward-only Excel
dotnet run -c Release --project benchmarks/DataExport.Benchmarks -- --filter '*ExcelExportBenchmarks*'
```

The separate **Streaming XLSX performance validation** GitHub Actions workflow
runs a 10k-row smoke check on relevant PRs and exposes manual 100k and 1M
stress options after the workflow is merged. It uploads JSON metrics
(elapsed time, allocations, process peak working set and file size).
A 1M-row run with a 500,000-row worksheet limit should produce three
sheets, each with its own header. The max-peak-mb flag makes a configured
memory budget enforceable; it is not a guarantee of the same peak on
different machines. Comparing runs requires the same runner hardware,
.NET runtime, data shape and workload. Peak *process* memory includes
runtime overhead and should not be confused with managed allocations.

### Measured XLSX stress results

GitHub Actions, Ubuntu runner, .NET 10.0.15, five fields per synthetic
record (ID, name, score, Boolean, date), using the forward-only exporter.
Each XLSX's worksheet row counts were checked with streaming XML parsing.
These figures include runtime and ZIP-package overhead, and should not be
generalized to all servers.

| Records | Time | Peak process RAM | Output size | Worksheet rows |
| ---: | ---: | ---: | ---: | --- |
| 100,000 | 3.385 s | 114.3 MB | 1.42 MB | 100,001 |
| 1,000,000 | 11.204 s | 486.68 MB | 14.37 MB | 500,000 + 500,000 + 3 |

**Baseline:** These measurements were taken using the previous Open XML SDK package implementation.

### Direct-ZIP writer results (November 22, 2026)

Same GitHub Actions performance harness, .NET 10.0.15, four reported
processors and five synthetic columns. XML worksheet row counts were
verified for each run.

| Records | Time | Peak process RAM | Output size |
| ---: | ---: | ---: | ---: |
| 100,000 | 1.527 s | 65.4 MB | 4.63 MB |
| 1,000,000 | 3.860 s | 134.09 MB | 46.47 MB |

Compared with the earlier Open XML package writer at one million rows,
peak working set fell from 486.68 MB to 134.09 MB (approximately 72% lower),
and elapsed export time fell from 11.204 s to 3.860 s. The trade-off:
the resulting XLSX grew from 14.37 MB to 46.47 MB because this
implementation favors faster ZIP compression. These are *observed*
results, not guaranteed limits across different input shapes or hosts.
The output was validated structurally and the small-file regression
tests opened it with ClosedXML.

**These numbers are controlled benchmark observations, not production guarantees.**
Re-run the benchmark with your own data shape, hosting environment, and throughput
constraints before advertising performance numbers. Generated XLSX files are
deleted automatically by the harness.

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
