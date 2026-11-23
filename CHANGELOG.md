# Changelog

## 2.0.0-rc.1 — Release candidate (not published; 2026-11-23)

This is a GitHub source candidate for a **major version 2.0.0** release.
It has **not** been published to NuGet. Stable release requires a separate,
manually approved 2.0.0 commit, final package checks and a guarded publishing action.

### Added
- Export plain C# objects, dictionaries and DataTables to JSON, XML, CSV and XLSX.
- Configurable CSV delimiter, headers, selected columns, formula protection and encoding.
- Async stream-based CSV/JSON/XML exports and export to writable caller-owned streams.
- Forward-only CSV export from `IAsyncEnumerable<T>`.
- Direct-ZIP streaming XLSX writer from async records or `DbDataReader`, supporting
  writable non-seekable streams, typed values and worksheet rollover.
- In-memory multi-sheet Excel report builder with mapped columns, styles and totals.
- Unit/integration tests, large-XLSX performance checks, dependency advisory CI,
  isolated packed-package consumers on .NET 8 and .NET 10, and guarded release workflow.

### Breaking changes / compatibility review
- `DataExporter` is now a `sealed partial class`; subclassing it is no longer supported.
- JSON serialization switched from Newtonsoft.Json to System.Text.Json. Some
  edge-case serialization behaviors may differ; test downstream applications.
- Existing `ToJson(string)`, `ToXml(string)`, `ToExcel(string)` and
  `ToCsv(string)` methods remain available; the legacy `ToExcel(path)`
  remains headerless and `ToCsv(path)` does not modify formula-like values.
- The old package description mentioned Word and PDF, but the historical
  source did not expose Word/PDF export methods. This release makes the
  supported-format list precise; Word/PDF remain unsupported by DataExport.
- The NuGet package now targets .NET 8 and .NET 10. The .NET 8 runtime passed
  its support end on November 10, 2026; use supported .NET 10 when possible.
- Streaming XLSX is structurally validated, not a general replacement for
  an Excel rendering engine; generated files have no charts or formulas.

### Performance context
- A GitHub-hosted .NET 10 benchmark (2026-11-22) measured a synthetic
  one-million-row XLSX export in 3.860 seconds at ~134.09 MB whole-process
  peak memory, producing ~46.47 MB of output. This is not a performance
  guarantee; compare workload shape, environment and generated workbook size.

### Release status
- Candidate source and CI are staged. NuGet publication, final release
  approval, a stable version tag, and GitHub Release remain pending.

## 1.2.0 development notes (incorporated into v2)

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
