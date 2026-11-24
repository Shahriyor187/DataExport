using System.Data.Common;
using System.Text;

namespace DataExportPackage;

public sealed partial class DataExporter
{
    /// <summary>
    /// Streams the current result set of a database reader to CSV. The reader must
    /// be positioned before its first record. Neither the reader nor output is disposed.
    /// Selected columns are read in source ordinal order for sequential-access readers.
    /// </summary>
    public static async Task WriteCsvAsync(DbDataReader reader, Stream output,
        DataExportOptions? options = null, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentNullException.ThrowIfNull(reader);
        ArgumentNullException.ThrowIfNull(output);
        if (!output.CanWrite)
            throw new ArgumentException("Output must be writable.", nameof(output));
        if (reader.IsClosed)
            throw new InvalidOperationException("Database reader is closed.");

        options ??= new DataExportOptions();
        ValidateCsvOptions(options);
        var progress = ExportProgressTracker.Create(options.Progress, options.ProgressIntervalRows, options.MaxRows == 0 ? 0 : options.TotalRows, cancellationToken);
        var names = Enumerable.Range(0, reader.FieldCount).Select(reader.GetName).ToArray();
        if (names.Any(string.IsNullOrWhiteSpace) ||
            names.Distinct(StringComparer.Ordinal).Count() != names.Length)
            throw new ArgumentException("Database field names must be unique and nonempty. Alias columns in the query.", nameof(reader));

        var columns = options.Columns?.ToArray() ?? names;
        if (columns.Any(string.IsNullOrWhiteSpace) ||
            columns.Distinct(StringComparer.Ordinal).Count() != columns.Length)
            throw new ArgumentException("Columns must contain unique, nonempty names.", nameof(options));
        var ordinals = names.Select((name, ordinal) => (name, ordinal))
            .ToDictionary(x => x.name, x => x.ordinal, StringComparer.Ordinal);
        var readOrder = new List<(int Source, int Target)>(columns.Length);
        for (var i = 0; i < columns.Length; i++)
        {
            if (!ordinals.TryGetValue(columns[i], out var ordinal))
                throw new ArgumentException($"Unknown database column '{columns[i]}'.", nameof(options));
            readOrder.Add((ordinal, i));
        }
        readOrder.Sort((a, b) => a.Source.CompareTo(b.Source));

        var delimiter = options.CsvDelimiter;
        var protect = options.CsvProtectFormulas;
        var remainingSkip = options.RowOffset;
        var limit = options.MaxRows;
        var fields = new string[columns.Length];
        cancellationToken.ThrowIfCancellationRequested();
        await using var writer = new StreamWriter(output, options.CsvEncoding ?? Utf8WithoutBom, 8192, leaveOpen: true);
        if (options.CsvIncludeHeaders && columns.Length > 0)
            await writer.WriteLineAsync(JoinCsv(columns.Select(c => CsvValue(c, protect)), delimiter).AsMemory(), cancellationToken).ConfigureAwait(false);

        long written = 0;
        while (limit is null || written < limit.Value)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var hasRow = await reader.ReadAsync(cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            if (!hasRow) break;
            if (remainingSkip > 0) { remainingSkip--; continue; }

            foreach (var (source, target) in readOrder)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var isNull = await reader.IsDBNullAsync(source, cancellationToken).ConfigureAwait(false);
                cancellationToken.ThrowIfCancellationRequested();
                fields[target] = isNull ? string.Empty : CsvValue(reader.GetValue(source), protect);
            }
            cancellationToken.ThrowIfCancellationRequested();
            await writer.WriteLineAsync(JoinCsv(fields, delimiter).AsMemory(), cancellationToken).ConfigureAwait(false);
            written++;
            progress?.RowWritten();
        }

        cancellationToken.ThrowIfCancellationRequested();
        await writer.FlushAsync(cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        progress?.Finish();
    }

    /// <summary>Streams database records to TSV with tab separators, leaving both caller-owned resources open.</summary>
    public static Task WriteTsvAsync(DbDataReader reader, Stream output,
        DataExportOptions? options = null, CancellationToken cancellationToken = default) =>
        WriteCsvAsync(reader, output, TsvOptions(options), cancellationToken);
}
