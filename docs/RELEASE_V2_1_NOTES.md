# DataExport v2.1.0 — release notes

Stable release candidate; publication is pending final validation and explicit approval.

## Added
- Dedicated TSV file, stream and async-record exports; ExportFormat.Tsv appended without renumbering previous values.
- DbDataReader-to-CSV/TSV forward-only exports, with schema discovery and selected-column ordering.
- RowOffset and MaxRows for buffered and streaming exports.
- Configurable row progress reporting, with optional estimated totals for streaming sources.

## Improved
- Database-to-XLSX export avoids retrieving fields from skipped records and unselected columns, and checks cancellation at provider boundaries.
- Cooperative cancellation across supported asynchronous export paths.
- Formula mitigation whitespace inspection avoids temporary trimmed strings.
- Documentation, production integration guidance and verified large-data benchmark tooling.

## Compatibility and security
- Targets net8.0 and net10.0; prefer a supported runtime for production.
- Legacy CSV and XLSX defaults are retained. Options-based CSV and new TSV APIs mitigate formula-like text by default.
- Caller-owned readers and streams stay open. Partial output must be discarded after failures.
- File exports are not transactional. Stream output does not automatically imply bounded memory; use forward-only APIs for large datasets.
- Progress counts processed data rows, not successful publication. Cancellation is cooperative, not a hard timeout.

## Release candidate validation (November 24, 2026)

- 106 regression cases passed on each of .NET 8 and .NET 10 (212 executions).
- One million synthetic records were exported and verified per format: CSV 1.149 s (66.51 MiB peak working set), TSV 1.068 s (63.11 MiB), XLSX 3.385 s (70.21 MiB). These are runner-specific measurements, not guaranteed production performance.
- No known vulnerabilities reported across 16 resolved dependency entries.
- Package validation against published v2.0.0 passed.
- Refreshed cross-package integration passed on .NET 8 and .NET 10. The final resource-boundary mutation was detected on both frameworks.
- Stable-version package validation remains a release gate. No package has been published.

See README and docs/PRODUCTION_GUIDE.md.
