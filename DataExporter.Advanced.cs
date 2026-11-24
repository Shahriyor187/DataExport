using ClosedXML.Excel;
using System.Data;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;

namespace DataExportPackage;

public sealed partial class DataExporter
{
    /// <summary>Creates an exporter from regular C# objects or dictionaries.</summary>
    public static DataExporter FromObjects<T>(IEnumerable<T> records)
    {
        ArgumentNullException.ThrowIfNull(records);
        var rows = new List<Dictionary<string, object>>();
        foreach (var record in records)
        {
            if (record is null)
                throw new ArgumentException("Records must not contain null objects.", nameof(records));

            rows.Add(ObjectToRow(record));
        }

        return new DataExporter(rows);
    }

    /// <summary>Creates an exporter from a System.Data.DataTable.</summary>
    public static DataExporter FromDataTable(DataTable table)
    {
        ArgumentNullException.ThrowIfNull(table);
        var rows = new List<Dictionary<string, object>>();
        foreach (DataRow row in table.Rows)
        {
            var data = new Dictionary<string, object>(StringComparer.Ordinal);
            foreach (DataColumn column in table.Columns)
                data[column.ColumnName] = row.IsNull(column) ? null! : row[column];
            rows.Add(data);
        }

        return new DataExporter(rows);
    }

    /// <summary>Exports C# objects directly to a file, without manually creating dictionaries.</summary>
    public static void Export<T>(IEnumerable<T> records, string filePath, ExportFormat format, DataExportOptions? options = null)
    {
        var exporter = FromObjects(records);
        switch (format)
        {
            case ExportFormat.Json:
                exporter.ToJson(filePath);
                break;
            case ExportFormat.Xml:
                exporter.ToXml(filePath);
                break;
            case ExportFormat.Csv:
                exporter.ToCsv(filePath, options);
                break;
            case ExportFormat.Excel:
                exporter.ToExcel(filePath, options);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(format));
        }
    }

    /// <summary>Writes exported records to a caller-owned stream, such as an HTTP response body.</summary>
    public async Task ExportAsync(Stream output, ExportFormat format, DataExportOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(output);
        cancellationToken.ThrowIfCancellationRequested();

        switch (format)
        {
            case ExportFormat.Json:
                await JsonSerializer.SerializeAsync(output, _data, JsonOptions, cancellationToken).ConfigureAwait(false);
                break;
            case ExportFormat.Xml:
                using (var writer = XmlWriter.Create(output, new XmlWriterSettings
                {
                    Async = true,
                    Indent = true,
                    Encoding = Utf8WithoutBom,
                    CloseOutput = false
                }))
                {
                    var root = new XElement("data");
                    foreach (var row in _data)
                    {
                        var entry = new XElement("entry");
                        foreach (var pair in row)
                        {
                            entry.Add(new XElement(XmlConvert.EncodeLocalName(string.IsNullOrEmpty(pair.Key) ? "_" : pair.Key),
                                FormatValue(pair.Value)));
                        }
                        root.Add(entry);
                    }
                    cancellationToken.ThrowIfCancellationRequested();
                    var document = new XDocument(new XDeclaration("1.0", "utf-8", null), root);
                    await document.SaveAsync(writer, cancellationToken).ConfigureAwait(false);
                }
                break;
            case ExportFormat.Csv:
                await ToCsvAsync(output, options, cancellationToken).ConfigureAwait(false);
                break;
            case ExportFormat.Excel:
                WriteExcelCore(output, options, cancellationToken);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(format));
        }
        cancellationToken.ThrowIfCancellationRequested();
    }

    /// <summary>Exports CSV without buffering the entire output in memory.</summary>
    public void ToCsv(string filePath, DataExportOptions? options)
    {
        ValidateFilePath(filePath);
        using var output = File.Create(filePath);
        using var writer = new StreamWriter(output, options?.CsvEncoding ?? Utf8WithoutBom, 8192, leaveOpen: true);
        WriteCsvRows(writer, options ?? new DataExportOptions());
    }

    /// <summary>Writes CSV to a caller-owned stream, leaving it open.</summary>
    public async Task ToCsvAsync(Stream output, DataExportOptions? options = null, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentNullException.ThrowIfNull(output);
        options ??= new DataExportOptions();
        ValidateCsvOptions(options);
        await using var writer = new StreamWriter(output, options.CsvEncoding ?? Utf8WithoutBom, 8192, leaveOpen: true);
        var columns = SelectColumns(options, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        if (options.CsvIncludeHeaders && columns.Count > 0)
            await writer.WriteLineAsync(JoinCsv(columns.Select(column => CsvValue(column, options.CsvProtectFormulas)), options.CsvDelimiter).AsMemory(), cancellationToken).ConfigureAwait(false);
        foreach (var row in _data)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var values = columns.Select(key =>
            {
                row.TryGetValue(key, out var value);
                return CsvValue(value, options.CsvProtectFormulas);
            });
            await writer.WriteLineAsync(JoinCsv(values, options.CsvDelimiter).AsMemory(), cancellationToken).ConfigureAwait(false);
        }
        cancellationToken.ThrowIfCancellationRequested();
        await writer.FlushAsync(cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
    }

    /// <summary>
    /// Streams CSV from an asynchronous source, without collecting all records in memory.
    /// Supply Columns for sparse dictionaries whose keys may differ between records.
    /// </summary>
    public static async Task WriteCsvAsync<T>(
        IAsyncEnumerable<T> records, Stream output, DataExportOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentNullException.ThrowIfNull(records);
        ArgumentNullException.ThrowIfNull(output);
        options ??= new DataExportOptions();
        ValidateCsvOptions(options);
        await using var writer = new StreamWriter(output, options.CsvEncoding ?? Utf8WithoutBom, 8192, leaveOpen: true);
        List<string>? columns = options.Columns?.ToList();
        var headerWritten = false;

        await foreach (var record in records.WithCancellation(cancellationToken).ConfigureAwait(false))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (record is null)
                throw new ArgumentException("Records must not contain null objects.", nameof(records));
            var row = ObjectToRow(record);
            columns ??= row.Keys.ToList();
            if (!headerWritten)
            {
                if (options.CsvIncludeHeaders && columns.Count > 0)
                    await writer.WriteLineAsync(JoinCsv(columns.Select(column => CsvValue(column, options.CsvProtectFormulas)), options.CsvDelimiter).AsMemory(), cancellationToken).ConfigureAwait(false);
                headerWritten = true;
            }

            var values = columns.Select(key =>
            {
                row.TryGetValue(key, out var value);
                return CsvValue(value, options.CsvProtectFormulas);
            });
            await writer.WriteLineAsync(JoinCsv(values, options.CsvDelimiter).AsMemory(), cancellationToken).ConfigureAwait(false);
        }
        cancellationToken.ThrowIfCancellationRequested();
        await writer.FlushAsync(cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
    }

    /// <summary>Writes a configurable Excel workbook to a file path.</summary>
    public void ToExcel(string filePath, DataExportOptions? options)
    {
        ValidateFilePath(filePath);
        using var output = File.Create(filePath);
        ToExcel(output, options);
    }

    /// <summary>Writes an Excel workbook to a caller-owned stream, preserving numeric, date and Boolean values.</summary>
    public void ToExcel(Stream output, DataExportOptions? options = null) =>
        WriteExcelCore(output, options, CancellationToken.None);

    private void WriteExcelCore(Stream output, DataExportOptions? options, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentNullException.ThrowIfNull(output);
        options ??= new DataExportOptions();
        ArgumentException.ThrowIfNullOrWhiteSpace(options.WorksheetName);
        var columns = SelectColumns(options, cancellationToken);
        var firstDataRow = options.ExcelIncludeHeaders && columns.Count > 0 ? 2 : 1;
        if ((long)_data.Count + firstDataRow - 1 > 1_048_576)
            throw new InvalidOperationException("The data exceeds Excel's worksheet row limit.");
        if (columns.Count > 16_384)
            throw new InvalidOperationException("The data exceeds Excel's worksheet column limit.");

        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add(options.WorksheetName);
        if (options.ExcelIncludeHeaders)
        {
            for (var index = 0; index < columns.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                worksheet.Cell(1, index + 1).Value = columns[index];
            }
            if (options.ExcelBoldHeaders && columns.Count > 0)
                worksheet.Row(1).Style.Font.Bold = true;
        }

        for (var rowIndex = 0; rowIndex < _data.Count; rowIndex++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var row = _data[rowIndex];
            for (var colIndex = 0; colIndex < columns.Count; colIndex++)
            {
                row.TryGetValue(columns[colIndex], out var value);
                SetCellValue(worksheet.Cell(rowIndex + firstDataRow, colIndex + 1), value);
            }
        }

        if (options.ExcelIncludeHeaders && columns.Count > 0)
        {
            if (options.ExcelFreezeHeader)
                worksheet.SheetView.FreezeRows(1);
            if (options.ExcelAutoFilter)
                worksheet.Range(1, 1, Math.Max(1, _data.Count + 1), columns.Count).SetAutoFilter();
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (columns.Count > 0 && options.ExcelAutoFitColumns)
            worksheet.Columns(1, columns.Count).AdjustToContents();
        cancellationToken.ThrowIfCancellationRequested();
        workbook.SaveAs(output);
        cancellationToken.ThrowIfCancellationRequested();
    }

    private static void SetCellValue(IXLCell cell, object? value)
    {
        switch (value)
        {
            case null or DBNull:
                cell.Value = "";
                break;
            case bool b:
                cell.Value = b;
                break;
            case byte or short or int or long or float or double or decimal:
                cell.Value = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                break;
            case DateTime date:
                cell.Value = date;
                break;
            case DateTimeOffset offset:
                cell.Value = offset.DateTime;
                break;
            case TimeSpan time:
                cell.Value = time;
                break;
            default:
                cell.Value = FormatValue(value);
                break;
        }
    }

    private List<string> SelectColumns(DataExportOptions options, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var columns = options.Columns?.ToList() ?? GetColumns(cancellationToken);
        if (columns.Any(string.IsNullOrWhiteSpace) || columns.Distinct(StringComparer.Ordinal).Count() != columns.Count)
            throw new ArgumentException("Columns must contain unique, nonempty names.", nameof(options));
        return columns;
    }

    private void WriteCsvRows(TextWriter writer, DataExportOptions options)
    {
        ValidateCsvOptions(options);
        var columns = SelectColumns(options);
        if (columns.Count == 0)
            return;

        if (options.CsvIncludeHeaders)
            writer.WriteLine(JoinCsv(columns.Select(column => CsvValue(column, options.CsvProtectFormulas)), options.CsvDelimiter));
        foreach (var row in _data)
        {
            var values = columns.Select(key =>
            {
                row.TryGetValue(key, out var value);
                return CsvValue(value, options.CsvProtectFormulas);
            });
            writer.WriteLine(JoinCsv(values, options.CsvDelimiter));
        }
    }

    private static void ValidateCsvOptions(DataExportOptions options)
    {
        if (options.CsvDelimiter is '\0' or '\r' or '\n' or '"')
            throw new ArgumentException("CSV delimiter must not be a quote or newline.", nameof(options));
    }

    private static string JoinCsv(IEnumerable<string> fields, char delimiter) =>
        string.Join(delimiter, fields.Select(value =>
        {
            if (value.IndexOfAny([delimiter, '"', '\r', '\n']) < 0)
                return value;
            return "\"" + value.Replace("\"", "\"\"", StringComparison.Ordinal) + "\"";
        }));

    private static string CsvValue(object? value, bool protectFormulas)
    {
        var formatted = FormatValue(value);
        if (protectFormulas && value is string && formatted.TrimStart().Length > 0 &&
            formatted.TrimStart()[0] is '=' or '+' or '-' or '@')
            return "'" + formatted;
        return formatted;
    }

    private static Dictionary<string, object> ObjectToRow<T>(T value)
    {
        if (value is IDictionary<string, object> dictionary)
            return new Dictionary<string, object>(dictionary, StringComparer.Ordinal);

        var properties = value!.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public)
            .Where(property => property.CanRead && property.GetIndexParameters().Length == 0);
        var row = new Dictionary<string, object>(StringComparer.Ordinal);
        foreach (var property in properties)
            row[property.Name] = property.GetValue(value)!;
        return row;
    }
}
