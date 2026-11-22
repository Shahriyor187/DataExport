using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data.Common;
using System.Globalization;
using System.Runtime.CompilerServices;

namespace DataExportPackage;

/// <summary>Options for the forward-only XLSX writer.</summary>
public sealed class StreamingExcelOptions
{
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
    /// Writes XLSX in constant working memory relative to row count using forward-only Open XML.
    /// Does not buffer the source or use ClosedXML's in-memory workbook.
    /// The caller owns and must dispose the destination stream.
    /// </summary>
    public static async Task WriteExcelStreamingAsync<T>(
        IAsyncEnumerable<T> records, Stream destination,
        StreamingExcelOptions? options = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(records);
        ValidateStreamingExcelArguments(destination, options ??= new StreamingExcelOptions());
        async IAsyncEnumerable<IReadOnlyDictionary<string, object?>> Rows(
            [EnumeratorCancellation] CancellationToken ct)
        {
            await foreach (var value in records.WithCancellation(ct).ConfigureAwait(false))
            {
                if (value is null)
                    throw new ArgumentException("Records cannot contain null objects.", nameof(records));
                yield return ObjectToRow(value);
            }
        }
        await WriteExcelStreamingCoreAsync(Rows(cancellationToken), destination, options, cancellationToken)
            .ConfigureAwait(false);
    }

    /// <summary>
    /// Writes a database query result straight to XLSX, advancing the supplied reader with ReadAsync.
    /// The reader and destination stream remain owned by the caller.
    /// </summary>
    public static async Task WriteExcelStreamingAsync(
        DbDataReader reader, Stream destination,
        StreamingExcelOptions? options = null, CancellationToken cancellationToken = default)
    {
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
                MaxRowsPerWorksheet = options.MaxRowsPerWorksheet
            };

        async IAsyncEnumerable<IReadOnlyDictionary<string, object?>> Rows(
            [EnumeratorCancellation] CancellationToken ct)
        {
            while (await reader.ReadAsync(ct).ConfigureAwait(false))
            {
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
    }

    private static async Task WriteExcelStreamingCoreAsync(
        IAsyncEnumerable<IReadOnlyDictionary<string, object?>> rows, Stream output,
        StreamingExcelOptions options, CancellationToken cancellationToken)
    {
        var columns = options.Columns?.ToList();
        ValidateStreamingColumns(columns);
        using var document = SpreadsheetDocument.Create(output, SpreadsheetDocumentType.Workbook, true);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        workbookPart.Workbook.AppendChild(new Sheets());
        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = CreateStreamingStyles();
        stylesPart.Stylesheet.Save();

        StreamingWorksheet? sheet = null;
        var sheetIndex = 0;
        try
        {
            await foreach (var values in rows.WithCancellation(cancellationToken).ConfigureAwait(false))
            {
                cancellationToken.ThrowIfCancellationRequested();
                columns ??= values.Keys.ToList();
                ValidateStreamingColumns(columns);
                if (sheet is null || sheet.WrittenRowCount >= options.MaxRowsPerWorksheet)
                {
                    if (sheet is not null)
                    {
                        if (!options.SplitLargeWorksheets)
                            throw new InvalidOperationException("The worksheet row limit was exceeded. Enable SplitLargeWorksheets.");
                        sheet.Dispose();
                    }
                    sheet = OpenStreamingSheet(workbookPart, ++sheetIndex, options, columns);
                }

                var cells = new object?[columns.Count];
                for (var i = 0; i < columns.Count; i++)
                {
                    values.TryGetValue(columns[i], out var value);
                    cells[i] = value;
                }
                sheet.WriteRow(cells, header: false);
            }

            // Empty sources still produce a valid workbook and, if specified, the header row.
            if (sheet is null)
                sheet = OpenStreamingSheet(workbookPart, ++sheetIndex, options, columns ?? []);
        }
        finally
        {
            sheet?.Dispose();
        }
        workbookPart.Workbook.Save();
    }

    private static void ValidateStreamingExcelArguments(Stream destination, StreamingExcelOptions options)
    {
        ArgumentNullException.ThrowIfNull(destination);
        // Open XML's ZIP packaging needs random-access read/write I/O even though
        // the worksheet XML itself is written forward-only.
        if (!destination.CanWrite || !destination.CanRead || !destination.CanSeek)
            throw new ArgumentException(
                "The destination must be a readable, writable, seekable stream (e.g. a read-write FileStream).",
                nameof(destination));
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

    private static StreamingWorksheet OpenStreamingSheet(
        WorkbookPart workbookPart, int index, StreamingExcelOptions options, IReadOnlyList<string> columns)
    {
        var suffix = index == 1 ? "" : $" ({index})";
        var name = options.WorksheetName[..Math.Min(options.WorksheetName.Length, 31 - suffix.Length)] + suffix;
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheets = workbookPart.Workbook.GetFirstChild<Sheets>()!;
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = (uint)index,
            Name = name
        });
        var writer = OpenXmlWriter.Create(worksheetPart);
        writer.WriteStartElement(new Worksheet());
        writer.WriteStartElement(new SheetData());
        var sheet = new StreamingWorksheet(writer);
        if (options.IncludeHeaders && columns.Count > 0)
        {
            sheet.WriteRow(columns.Select(column =>
            {
                if (options.ColumnHeaders?.TryGetValue(column, out var label) == true)
                    return (object?)label;
                return column;
            }).ToArray(), header: true);
        }
        return sheet;
    }

    private static Stylesheet CreateStreamingStyles() => new(
        new Fonts(new Font(), new Font(new Bold())),
        new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }),
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 })),
        new Borders(new Border()),
        new CellStyleFormats(new CellFormat()),
        new CellFormats(
            new CellFormat(),
            new CellFormat { FontId = 1U, ApplyFont = true },
            new CellFormat { NumberFormatId = 22U, ApplyNumberFormat = true }));

    private sealed class StreamingWorksheet(OpenXmlWriter writer) : IDisposable
    {
        private bool _disposed;
        public int WrittenRowCount { get; private set; }

        public void WriteRow(IReadOnlyList<object?> values, bool header)
        {
            writer.WriteStartElement(new Row { RowIndex = (uint)++WrittenRowCount });
            foreach (var value in values)
                writer.WriteElement(ToOpenXmlCell(value, header));
            writer.WriteEndElement();
        }

        private static Cell ToOpenXmlCell(object? value, bool header)
        {
            var cell = new Cell();
            if (header) cell.StyleIndex = 1U;
            if (value is null or DBNull) return cell;
            if (value is bool boolean)
            {
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new CellValue(boolean ? "1" : "0");
            }
            else if (value is DateTime date && date.Year >= 1900)
            {
                cell.StyleIndex = 2U;
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(date.ToOADate().ToString("R", CultureInfo.InvariantCulture));
            }
            else if (value is byte or sbyte or short or ushort or int or uint or float or double or decimal ||
                     value is long longValue && longValue is >= -999999999999999 and <= 999999999999999 ||
                     value is ulong ulongValue && ulongValue <= 999999999999999)
            {
                cell.DataType = CellValues.Number;
                cell.CellValue = new CellValue(Convert.ToString(value, CultureInfo.InvariantCulture));
            }
            else
            {
                cell.DataType = CellValues.InlineString;
                cell.InlineString = new InlineString(
                    new Text(Convert.ToString(value, CultureInfo.InvariantCulture) ?? "")
                    { Space = SpaceProcessingModeValues.Preserve });
            }
            return cell;
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            writer.WriteEndElement(); // SheetData
            writer.WriteEndElement(); // Worksheet
            writer.Dispose();
        }
    }
}
