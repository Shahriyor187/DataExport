# Changelog

## 1.2.0 — Unreleased

### Added
- Export regular C# objects and DataTables without manual dictionary conversion.
- Stream-based JSON, XML, CSV and Excel output, including asynchronous CSV generation from `IAsyncEnumerable<T>`.
- Configurable CSV delimiter, column selection, and optional formula-injection protection.
- Excel worksheet naming, header styling, filtering, frozen rows, and typed numeric/date/Boolean cells.
- ASP.NET Core response-stream example and expanded test coverage.

### Compatibility
- Existing export method signatures remain intact.
- The original headerless `ToExcel(path)` behavior is preserved.
- The original `ToCsv(path)` output remains unmodified; use the new options-based API for formula protection.

## 1.1.0 — 2026-11-22

- Modernized .NET 8/.NET 10 packaging, dependencies, documentation and CI.
- Added CSV export, improved XML formatting and Excel column alignment.
