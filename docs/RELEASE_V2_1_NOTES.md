# DataExport v2.1.0 — draft release notes

Not published. Final release depends on Phase 1.8 verification.

## Added
- Dedicated TSV file, stream and async-record exports; ExportFormat.Tsv appended without renumbering previous values.
- DbDataReader-to-CSV/TSV forward-only exports, with schema discovery and selected-column ordering.
- RowOffset and MaxRows for buffered and streaming exports.
- Configurable row progress reporting, with optional estimated totals for streaming sources.

## Improved
- Cooperative cancellation across supported asynchronous export paths.
- Formula mitigation whitespace inspection avoids temporary trimmed strings.
- Documentation, production integration guidance and verified large-data benchmark tooling.

## Compatibility and security
- Targets net8.0 and net10.0; prefer a supported runtime for production.
- Legacy CSV and XLSX defaults are retained. Options-based CSV and new TSV APIs mitigate formula-like text by default.
- Caller-owned readers and streams stay open. Partial output must be discarded after failures.
- File exports are not transactional. Stream output does not automatically imply bounded memory; use forward-only APIs for large datasets.
- Progress counts processed data rows, not successful publication. Cancellation is cooperative, not a hard timeout.

See README and docs/PRODUCTION_GUIDE.md. Final test counts, benchmark observations and dependency audit results will be recorded after release validation.
