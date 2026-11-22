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
| `ToExcel(path)` | XLSX | Creates a worksheet named `Data` and includes a header row. |

CSV and Excel use the union of all dictionary keys, in the order those keys are first encountered. If a record does not contain a particular key, the corresponding cell is empty.

To omit the Excel header row:

```csharp
exporter.ToExcel("people-without-headers.xlsx", includeHeaders: false);
```

## Notes

- The library writes to the file path you provide. The destination directory must already exist.
- Values in CSV, XML, and Excel are formatted using invariant culture where applicable.
- Null values are exported as empty strings in CSV, XML, and Excel. JSON preserves null values.
- The package does not currently export Word or PDF files.

## License

MIT. See [LICENSE](LICENSE).
