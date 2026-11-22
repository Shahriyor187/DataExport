using ClosedXML.Excel;
using System.Data;
using System.Globalization;
using System.Reflection;

namespace DataExportPackage;

/// <summary>Defines a source column and its display behavior in an Excel report.</summary>
public sealed record ExcelReportColumn(
    string PropertyPath,
    string Header,
    string? NumberFormat = null,
    double? Width = null);

/// <summary>Configuration for one report worksheet.</summary>
public sealed class ExcelReportSheetOptions
{
    /// <summary>Optional property paths and display names. Nested paths such as Address.City are supported.</summary>
    public IReadOnlyList<ExcelReportColumn>? Columns { get; set; }
    public bool FreezeHeaders { get; set; } = true;
    public bool AutoFilter { get; set; } = true;
    public bool AutoFitColumns { get; set; } = true;
    public bool AlternatingRowColors { get; set; } = true;
}

/// <summary>
/// Creates multi-sheet XLSX reports with typed cells, custom columns and optional totals.
/// Unlike WriteExcelStreamingAsync, this builder uses an in-memory ClosedXML workbook.
/// </summary>
public sealed class ExcelReportBuilder : IDisposable
{
    private readonly XLWorkbook _workbook = new();
    private readonly Dictionary<string, (int Columns, int LastDataRow)> _sheets = new(StringComparer.OrdinalIgnoreCase);
    private bool _disposed;

    public ExcelReportBuilder AddSheet<T>(
        string sheetName, IEnumerable<T> records, ExcelReportSheetOptions? options = null)
    {
        ThrowIfDisposed();
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentNullException.ThrowIfNull(records);
        options ??= new ExcelReportSheetOptions();

        var rowList = records.Select(ToDictionary).ToList();
        var columns = options.Columns?.ToList()
            ?? rowList.SelectMany(row => row.Keys).Distinct(StringComparer.Ordinal)
                .Select(key => new ExcelReportColumn(key, key)).ToList();

        if (columns.Count > 16_384 || columns.Any(c => string.IsNullOrWhiteSpace(c.PropertyPath) ||
             string.IsNullOrWhiteSpace(c.Header)) ||
            columns.Select(c => c.PropertyPath).Distinct(StringComparer.Ordinal).Count() != columns.Count)
            throw new ArgumentException("Invalid report columns.", nameof(options));
        if (rowList.Count > 1_048_575)
            throw new InvalidOperationException("Report sheet exceeds Excel's row limit.");

        var sheet = _workbook.Worksheets.Add(sheetName);
        for (var c = 0; c < columns.Count; c++)
        {
            var column = columns[c];
            sheet.Cell(1, c + 1).Value = column.Header;
            if (column.NumberFormat is not null)
                sheet.Column(c + 1).Style.NumberFormat.Format = column.NumberFormat;
            if (column.Width is double width)
            {
                if (width is <= 0 or > 255) throw new ArgumentOutOfRangeException(nameof(options));
                sheet.Column(c + 1).Width = width;
            }
        }
        if (columns.Count > 0)
        {
            var heading = sheet.Range(1, 1, 1, columns.Count);
            heading.Style.Font.Bold = true;
            heading.Style.Font.FontColor = XLColor.White;
            heading.Style.Fill.BackgroundColor = XLColor.FromHtml("#24466B");
        }

        for (var r = 0; r < rowList.Count; r++)
        {
            var source = rowList[r];
            for (var c = 0; c < columns.Count; c++)
            {
                if (!TryGetPath(source, columns[c].PropertyPath, out var value))
                    continue;
                WriteTypedCell(sheet.Cell(r + 2, c + 1), value);
            }
            if (options.AlternatingRowColors && r % 2 == 1 && columns.Count > 0)
                sheet.Range(r + 2, 1, r + 2, columns.Count).Style.Fill.BackgroundColor =
                    XLColor.FromHtml("#F1F5F9");
        }
        if (columns.Count > 0)
        {
            if (options.FreezeHeaders) sheet.SheetView.FreezeRows(1);
            if (options.AutoFilter)
                sheet.Range(1, 1, Math.Max(2, rowList.Count + 1), columns.Count).SetAutoFilter();
            if (options.AutoFitColumns)
            {
                for (var c = 0; c < columns.Count; c++)
                    if (columns[c].Width is null)
                        sheet.Column(c + 1).AdjustToContents();
            }
        }

        _sheets.Add(sheetName, (columns.Count, rowList.Count + 1));
        return this;
    }

    /// <summary>Adds a bold total row with SUM formulas for the specified 1-based column indices.</summary>
    public ExcelReportBuilder AddTotals(string sheetName, params int[] sumColumns)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(sumColumns);
        if (!_sheets.TryGetValue(sheetName, out var info))
            throw new ArgumentException("The worksheet was not created by this builder.", nameof(sheetName));
        if (info.Columns == 0)
            throw new InvalidOperationException("Cannot total a worksheet without columns.");

        var sheet = _workbook.Worksheet(sheetName);
        var totalsRow = info.LastDataRow + 1;
        sheet.Cell(totalsRow, 1).Value = "Total";
        foreach (var column in sumColumns.Distinct())
        {
            if (column < 1 || column > info.Columns)
                throw new ArgumentOutOfRangeException(nameof(sumColumns));
            if (column == 1)
                throw new ArgumentException("Column 1 is reserved for the Total label.", nameof(sumColumns));
            var columnLetter = XLHelper.GetColumnLetterFromNumber(column);
            sheet.Cell(totalsRow, column).FormulaA1 = $"SUM({columnLetter}2:{columnLetter}{info.LastDataRow})";
        }
        sheet.Range(totalsRow, 1, totalsRow, info.Columns).Style.Font.Bold = true;
        sheet.Range(totalsRow, 1, totalsRow, info.Columns).Style.Fill.BackgroundColor =
            XLColor.FromHtml("#E1EDF8");
        return this;
    }

    public void SaveAs(string path)
    {
        ThrowIfDisposed();
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        _workbook.SaveAs(path);
    }

    public void SaveAs(Stream destination)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(destination);
        _workbook.SaveAs(destination);
    }

    private static Dictionary<string, object?> ToDictionary<T>(T record)
    {
        if (record is null) throw new ArgumentException("Records must not contain null values.");
        if (record is DataRow dataRow)
            return dataRow.Table.Columns.Cast<DataColumn>()
                .ToDictionary(c => c.ColumnName, c => dataRow.IsNull(c) ? null : dataRow[c],
                    StringComparer.Ordinal);
        if (record is IDictionary<string, object> dictionary)
            return dictionary.ToDictionary(x => x.Key, x => (object?)x.Value, StringComparer.Ordinal);
        return record.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public)
            .Where(p => p.CanRead && p.GetIndexParameters().Length == 0)
            .ToDictionary(p => p.Name, p => p.GetValue(record), StringComparer.Ordinal);
    }

    private static bool TryGetPath(
        IReadOnlyDictionary<string, object?> source, string propertyPath, out object? value)
    {
        value = source;
        foreach (var part in propertyPath.Split('.'))
        {
            if (value is IReadOnlyDictionary<string, object?> readOnly && readOnly.TryGetValue(part, out var next))
                value = next;
            else if (value is IDictionary<string, object> dictionary && dictionary.TryGetValue(part, out var nested))
                value = nested;
            else if (value is not null)
            {
                var property = value.GetType().GetProperty(part, BindingFlags.Instance | BindingFlags.Public);
                if (property is null || !property.CanRead || property.GetIndexParameters().Length != 0)
                {
                    value = null;
                    return false;
                }
                value = property.GetValue(value);
            }
            else return true;
        }
        return true;
    }

    private static void WriteTypedCell(IXLCell cell, object? value)
    {
        switch (value)
        {
            case null or DBNull: break;
            case bool boolean: cell.Value = boolean; break;
            case DateTime date: cell.Value = date; break;
            case DateTimeOffset offset: cell.Value = offset.DateTime; break;
            case TimeSpan span: cell.Value = span; break;
            case byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal:
                cell.Value = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                break;
            default: cell.Value = Convert.ToString(value, CultureInfo.InvariantCulture) ?? ""; break;
        }
    }

    private void ThrowIfDisposed()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _workbook.Dispose();
    }
}
