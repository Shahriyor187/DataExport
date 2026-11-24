# Changelog

## Unreleased (v2.1 development)

- Improve cooperative cancellation in existing async CSV, XML, buffered XLSX and streaming XLSX exports without changing public signatures. Document synchronous library boundaries and partial-stream ownership.

## 2.0.0 — Stable source release prepared (2026-11-23; publication pending)

This release contains substantial new export features and intentional
compatibility changes. **Source code is ready for a public 2.0.0 package,
but NuGet publication requires a separate manually approved workflow.**

### New in v2
- Export C# objects, dictionaries and DataTables to JSON, XML, CSV and XLSX.
- Configure CSV headers/delimiters, selected columns, encodings and formula protection.
- Stream outputs to caller-owned streams, including asynchronous CSV from
  IAsyncEnumerable sources.
- Write XLSX incrementally via a direct-ZIP writer using async records
  or DbDataReader, including worksheet rollover and typed values.
- Build styled multi-sheet Excel reports with custom columns and totals.
- Automated testing, isolated local .NET 8/.NET 10 package consumers,
  direct/transitive NuGet advisory audits and large-data XLSX benchmarks.

### Upgrade and compatibility notes
- DataExporter is now sealed; subclasses must be refactored.
- JSON serialization switched from Newtonsoft.Json to System.Text.Json;
  consumers sensitive to exact JSON formatting should regression-test.
- Legacy ToJson/ToXml/ToExcel/ToCsv method signatures remain available.
  The original ToExcel(path) still produces headerless output, and the
  original ToCsv(path) does not enable formula protection.
- The original package description referred to Word/PDF, but historical
  source did not implement those export methods; DataExport supports
  JSON, XML, CSV and Excel only.
- Targets net8.0 and net10.0; .NET 8 runtime support ended 2026-11-10.
  Prefer supported .NET 10 for production deployments.

### Validation evidence
- Source RC passed 28 tests per framework, two independently executable
  local package consumers and 10K-row streaming XLSX smoke test.
- The 2026-11-22 synthetic one-million-row benchmark observed 3.860 s,
  ~134.09 MB process peak and ~46.47 MB XLSX output. These values are
  test-run observations, not production guarantees.
- The final stable package will be validated again on the exact
  release commit before publication. NuGet publication and GitHub
  release/tag are pending approval.

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
