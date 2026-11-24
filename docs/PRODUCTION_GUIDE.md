# Production integration guide (v2.1 candidate)

These examples require the v2.1 source or the corresponding released package. Prefer a supported .NET 10 runtime.

## Authenticated database download

The application must register a provider-specific DbDataSource, configure authentication/authorization, and enforce its own tenant-level access policy. This example assumes the policy authorizes the complete report. Do not reuse this query for a multi-tenant report without an authorization filter.

```csharp
using System.Data;
using System.Data.Common;
using DataExportPackage;

app.MapGet("/reports/orders.csv", async (
    HttpContext context, DbDataSource database) =>
{
    var ct = context.RequestAborted;
    await using var connection = await database.OpenConnectionAsync(ct);
    await using var command = connection.CreateCommand();
    command.CommandText = "SELECT id AS Id, customer AS Customer FROM orders ORDER BY id";
    command.CommandTimeout = 30;
    await using var reader = await command.ExecuteReaderAsync(
        CommandBehavior.SequentialAccess, ct);

    context.Response.ContentType = "text/csv; charset=utf-8";
    context.Response.Headers.ContentDisposition = "attachment; filename=orders.csv";
    context.Response.Headers.CacheControl = "no-store";

    await DataExporter.WriteCsvAsync(reader, context.Response.Body,
        new DataExportOptions
        {
            CsvProtectFormulas = true,
            MaxRows = 100_000
        }, ct);
}).RequireAuthorization("Reports.ReadAll");
```

Use provider-appropriate quoting for case-sensitive aliases. Parameterize all user-controlled query values. Apply database-side authorization, pagination and resource limits before exporting. MaxRows bounds emitted records, not database query cost. The caller owns the reader, connection and output stream. The library does not execute SQL or enforce access control. On a failure after HTTP headers have been sent, do not append an error document to the CSV response; abort the response and treat it as incomplete.

## Memory and output safety

Use the async-record or database streaming APIs for large result sets. FromObjects, FromDataTable, the buffered workbook and report builder retain input in memory. Individual large database fields and provider buffers can still consume considerable memory. A non-seekable output stream is supported by the forward-only writers. Limit concurrent exports, query duration, row counts, individual field sizes and total output/storage at the application boundary.

Options-based CSV and TSV default to spreadsheet formula protection. The legacy single-argument CSV API retains its historical behavior; use the options-based API for untrusted data. Apostrophe prefixing is a mitigation, not a guarantee for every spreadsheet program. TSV uses CSV-style quoting, so use a compatible parser. Do not create formulas from untrusted strings in report-builder APIs.

DataExport file writes are not transactional. For reliable publication, write to an application-owned private staging directory on the destination filesystem, close the output, await success, then rename the file. Clean up in a finally block. Never use an uploaded filename or untrusted path directly. Avoid overwriting existing files until conversion succeeds.

Cancellation is cooperative. A synchronous serializer, workbook save or provider call may not stop immediately. Discard partial output on failure. Progress rows are not a success signal; await the export. Keep callbacks fast and avoid mutating options or resources concurrently. Do not log sensitive row values.

## Reproducing progress overhead

```bash
dotnet run -c Release --project benchmarks/DataExport.Benchmarks -- --progress
```

See README for the separate XLSX stress runner and the historical measurements. Measure on representative data and target infrastructure; process peak memory is not total allocated memory.
