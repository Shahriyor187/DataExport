namespace DataExportPackage;

public sealed partial class DataExporter
{
    /// <summary>Writes tab-separated data to a file using CSV-style quoting.</summary>
    public void ToTsv(string filePath, DataExportOptions? options = null) =>
        ToCsv(filePath, TsvOptions(options));

    /// <summary>Writes tab-separated data to a caller-owned stream, leaving it open.</summary>
    public Task ToTsvAsync(Stream output, DataExportOptions? options = null,
        CancellationToken cancellationToken = default) =>
        ToCsvAsync(output, TsvOptions(options), cancellationToken);

    /// <summary>Streams tab-separated records without buffering the entire source.</summary>
    public static Task WriteTsvAsync<T>(IAsyncEnumerable<T> records, Stream output,
        DataExportOptions? options = null, CancellationToken cancellationToken = default) =>
        WriteCsvAsync(records, output, TsvOptions(options), cancellationToken);

    private static DataExportOptions TsvOptions(DataExportOptions? options)
    {
        options ??= new DataExportOptions();
        return new DataExportOptions
        {
            Columns = options.Columns,
            RowOffset = options.RowOffset,
            MaxRows = options.MaxRows,
            CsvDelimiter = ',',
            CsvIncludeHeaders = options.CsvIncludeHeaders,
            CsvEncoding = options.CsvEncoding,
            CsvProtectFormulas = options.CsvProtectFormulas
        };
    }
}
