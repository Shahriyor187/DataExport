using ClosedXML.Excel;
using System.Globalization;
using System.Text;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;

namespace DataExportPackage;

/// <summary>
/// Exports a collection of dictionary records to JSON, XML, CSV, or Excel (XLSX).
/// </summary>
/// <remarks>
/// Excel and CSV exports use the union of keys found across all records, preserving
/// the order in which keys are first encountered. Excel and CSV exports include a
/// header row by default.
/// </remarks>
public sealed partial class DataExporter
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true
    };

    private readonly IReadOnlyList<Dictionary<string, object>> _data;

    /// <summary>
    /// Creates an exporter for the supplied records.
    /// </summary>
    /// <param name="data">Records to export. Each dictionary represents one row.</param>
    /// <exception cref="ArgumentNullException">The data collection is null.</exception>
    /// <exception cref="ArgumentException">The collection contains a null record.</exception>
    public DataExporter(List<Dictionary<string, object>> data)
    {
        ArgumentNullException.ThrowIfNull(data);

        if (data.Any(record => record is null))
        {
            throw new ArgumentException("The data collection cannot contain null records.", nameof(data));
        }

        _data = data;
    }

    /// <summary>
    /// Writes the records as indented JSON using UTF-8 without a byte-order mark.
    /// </summary>
    /// <param name="filePath">Destination file path.</param>
    public void ToJson(string filePath)
    {
        ValidateFilePath(filePath);
        var json = JsonSerializer.Serialize(_data, JsonOptions);
        File.WriteAllText(filePath, json, Utf8WithoutBom);
    }

    /// <summary>
    /// Writes the records as XML. Dictionary keys are encoded as valid XML element names.
    /// </summary>
    /// <param name="filePath">Destination file path.</param>
    public void ToXml(string filePath)
    {
        ValidateFilePath(filePath);

        var document = new XDocument(
            new XDeclaration("1.0", "utf-8", null),
            new XElement("data",
                _data.Select(record =>
                    new XElement("entry",
                        record.Select(pair =>
                            new XElement(
                                XmlConvert.EncodeLocalName(string.IsNullOrEmpty(pair.Key) ? "_" : pair.Key),
                                FormatValue(pair.Value)))))));

        using var writer = XmlWriter.Create(filePath, new XmlWriterSettings
        {
            Encoding = Utf8WithoutBom,
            Indent = true
        });

        document.Save(writer);
    }

    /// <summary>
    /// Writes the records as an Excel workbook (.xlsx), without a header row by default.
    /// </summary>
    /// <param name="filePath">Destination file path.</param>
    public void ToExcel(string filePath) => ToExcel(filePath, includeHeaders: false);

    /// <summary>
    /// Writes the records as an Excel workbook (.xlsx).
    /// </summary>
    /// <param name="filePath">Destination file path.</param>
    /// <param name="includeHeaders">Whether to write the dictionary keys as the first row.</param>
    public void ToExcel(string filePath, bool includeHeaders) =>
        ToExcel(filePath, new DataExportOptions { ExcelIncludeHeaders = includeHeaders });

    /// <summary>
    /// Writes the records as UTF-8 CSV using RFC 4180-style escaping and a header row.
    /// </summary>
    /// <param name="filePath">Destination file path.</param>
    public void ToCsv(string filePath) => ToCsv(filePath, options: null);

    private List<string> GetColumns()
    {
        var columns = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);

        foreach (var record in _data)
        {
            foreach (var key in record.Keys)
            {
                if (seen.Add(key))
                {
                    columns.Add(key);
                }
            }
        }

        return columns;
    }

    private static void AppendCsvRecord(StringBuilder builder, IEnumerable<string> values)
    {
        builder.AppendJoin(',', values.Select(EscapeCsvValue));
        builder.AppendLine();
    }

    private static string EscapeCsvValue(string value)
    {
        if (value.IndexOfAny([',', '"', '\r', '\n']) < 0)
        {
            return value;
        }

        return $"\"{value.Replace("\"", "\"\"", StringComparison.Ordinal)}\"";
    }

    private static string FormatValue(object? value)
    {
        return value switch
        {
            null => string.Empty,
            IFormattable formattable => formattable.ToString(null, CultureInfo.InvariantCulture) ?? string.Empty,
            _ => value.ToString() ?? string.Empty
        };
    }

    private static void ValidateFilePath(string filePath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
    }

    private static readonly Encoding Utf8WithoutBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
}
