using System.Data.Common;
using System.Runtime.CompilerServices;

namespace DataExportPackage;

/// <summary>Options for the forward-only XLSX writer.</summary>
public sealed class StreamingExcelOptions
{
    /// <summary>Optional row progress notifications. Await the export itself to determine success.</summary>
    public IProgress<ExportProgress>? Progress { get; init; }
    /// <summary>Number of data rows between progress notifications.</summary>
    public int ProgressIntervalRows { get; init; } = 1000;
    /// <summary>Optional estimated number of output data rows for streaming sources, after row selection.</summary>
    public long? TotalRows { get; init; }

    /// <summary>Number of data records to skip before exporting. Headers are excluded.</summary>
    public long RowOffset { get; init; }
    /// <summary>Maximum data records to export, or null for unlimited. Zero exports no data.</summary>
    public long? MaxRows { get; init; }

    /// <summary>Explicit input property/field names and order. Required when later records have different keys.</summary>
    public IReadOnlyList<string>? Columns { get; set; }

    /// <summary>Optional display headers keyed by input column name.</summary>
    public IReadOnlyDictionary<string, string>? ColumnHeaders { get; set; }

    /// <summary>Base worksheet name; automatically suffixed for additional sheets.</summary>
    public string WorksheetName { get; set; } = "Data";

    /// <summary>Include a bold header row on every worksheet.</summary>
    public bool IncludeHeaders { get; set; } = true;

    /// <summary>Split into additional worksheets when the row limit is reached.</summary>
    public bool SplitLargeWorksheets { get; set; } = true;

    /// <summary>Maximum physical rows per sheet, including any header. Default is Excel's actual limit.</summary>
    public int MaxRowsPerWorksheet { get; set; } = 1_048_576;
}

public sealed partial class DataExporter
{
    /// <summary>
    /// Writes worksheet XML incrementally from the source instead of buffering all input rows.
    /// Open XML package generation may still increase process memory as output grows.
    /// The caller owns and must dispose the destination stream.
    /// </summary>
    public static async Task WriteExcelStreamingAsync<T>(
        IAsyncEnumerable<T> records, Stream destination,
        StreamingExcelOptions? options = null, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentNullException.ThrowIfNull(records);
        ValidateStreamingExcelArguments(destination, options ??= new StreamingExcelOptions());
        async IAsyncEnumerable<IReadOnlyDictionary<string, object?>> Rows(
            [EnumeratorCancellation] CancellationToken ct)
        {
            await foreach (var value in records.WithCancellation(ct).ConfigureAwait(false))
            {
                ct.ThrowIfCancellationRequested();
                if (value is null)
                    throw new ArgumentException("Records cannot contain null objects.", nameof(records));
                yield return ObjectToRow(value);
            }
        }
        await WriteExcelStreamingCoreAsync(Rows(cancellationToken), destination, options, cancellationToken)
            .ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
    }

    /// <summary>
    /// Writes a database query result straight to XLSX, advancing the supplied reader with ReadAsync.
    /// The reader and destination stream remain owned by the caller.
    /// </summary>
    public static async Task WriteExcelStreamingAsync(
        DbDataReader reader, Stream destination,
        StreamingExcelOptions? options = null, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentNullException.ThrowIfNull(reader);
        ValidateStreamingExcelArguments(destination, options ??= new StreamingExcelOptions());
        var fieldNames = Enumerable.Range(0, reader.FieldCount).Select(reader.GetName).ToArray();
        if (fieldNames.Distinct(StringComparer.Ordinal).Count() != fieldNames.Length)
            throw new ArgumentException("Database field names must be unique. Alias duplicate columns in the query.", nameof(reader));
        if (options.Columns is null)
            options = new StreamingExcelOptions
            {
                Columns = fieldNames,
                ColumnHeaders = options.ColumnHeaders,
                WorksheetName = options.WorksheetName,
                IncludeHeaders = options.IncludeHeaders,
                SplitLargeWorksheets = options.SplitLargeWorksheets,
                MaxRowsPerWorksheet = options.MaxRowsPerWorksheet,
                RowOffset = options.RowOffset,
                MaxRows = options.MaxRows,
                Progress = options.Progress,
                ProgressIntervalRows = options.ProgressIntervalRows,
                TotalRows = options.TotalRows
            };

        async IAsyncEnumerable<IReadOnlyDictionary<string, object?>> Rows(
            [EnumeratorCancellation] CancellationToken ct)
        {
            while (await reader.ReadAsync(ct).ConfigureAwait(false))
            {
                ct.ThrowIfCancellationRequested();
                var row = new Dictionary<string, object?>(fieldNames.Length, StringComparer.Ordinal);
                for (var i = 0; i < fieldNames.Length; i++)
                    row.Add(fieldNames[i], await reader.IsDBNullAsync(i, ct).ConfigureAwait(false)
                        ? null
                        : reader.GetValue(i));
                yield return row;
            }
        }

        await WriteExcelStreamingCoreAsync(Rows(cancellationToken), destination, options, cancellationToken)
            .ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
    }

    private static Task WriteExcelStreamingCoreAsync(
        IAsyncEnumerable<IReadOnlyDictionary<string, object?>> rows, Stream output,
        StreamingExcelOptions options, CancellationToken cancellationToken) =>
        WriteExcelZipCoreAsync(rows, output, options, cancellationToken);

    private static void ValidateStreamingExcelArguments(Stream destination, StreamingExcelOptions options)
    {
        ValidateRowWindow(options.RowOffset, options.MaxRows);
        ExportProgressTracker.Validate(options.ProgressIntervalRows, options.TotalRows);
        ArgumentNullException.ThrowIfNull(destination);
        // ZipArchiveMode.Create supports writable non-seekable streams.
        if (!destination.CanWrite)
            throw new ArgumentException("Destination must be a writable stream.", nameof(destination));
        if (options.MaxRowsPerWorksheet < (options.IncludeHeaders ? 2 : 1) ||
            options.MaxRowsPerWorksheet > 1_048_576)
            throw new ArgumentOutOfRangeException(nameof(options.MaxRowsPerWorksheet));
        ArgumentException.ThrowIfNullOrWhiteSpace(options.WorksheetName);
        if (options.WorksheetName.IndexOfAny(['[', ']', ':', '*', '?', '/', '\\']) >= 0 ||
            options.WorksheetName.Length > 31)
            throw new ArgumentException("Invalid Excel worksheet name.", nameof(options));
        ValidateStreamingColumns(options.Columns);
    }

    private static void ValidateStreamingColumns(IReadOnlyList<string>? columns)
    {
        if (columns is null) return;
        if (columns.Count > 16_384)
            throw new ArgumentException("Excel supports at most 16,384 columns.", nameof(columns));
        if (columns.Any(string.IsNullOrWhiteSpace) ||
            columns.Distinct(StringComparer.Ordinal).Count() != columns.Count)
            throw new ArgumentException("Column names must be unique and nonempty.", nameof(columns));
    }

}
