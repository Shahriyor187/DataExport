using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace DataExportPackage;

public sealed partial class DataExporter
{
    private const string SpreadsheetNamespace =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string OfficeRelationshipNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelationshipNamespace =
        "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string ContentTypeNamespace =
        "http://schemas.openxmlformats.org/package/2006/content-types";
    private const string OfficeRelationshipBase =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    private const string XmlNamespace = "http://www.w3.org/XML/1998/namespace";

    /// <summary>
    /// Creates a valid XLSX Open Packaging Conventions archive directly in Create mode.
    /// ZIP central-directory bookkeeping scales with sheet count, not input row count.
    /// This supports writable, non-seekable destinations such as HTTP response streams.
    /// </summary>
    private static async Task WriteExcelZipCoreAsync(
        IAsyncEnumerable<IReadOnlyDictionary<string, object?>> records,
        Stream destination, StreamingExcelOptions options, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var columns = options.Columns?.ToList();
        ValidateStreamingColumns(columns);
        var sheetNames = new List<string>();
        using var zip = new ZipArchive(destination, ZipArchiveMode.Create, leaveOpen: true);

        StreamingZipWorksheet? sheet = null;
        long remainingSkip = options.RowOffset, written = 0;
        try
        {
            if (options.MaxRows != 0)
            await foreach (var row in records.WithCancellation(cancellationToken).ConfigureAwait(false))
            {
                cancellationToken.ThrowIfCancellationRequested();
                columns ??= row.Keys.ToList();
                ValidateStreamingColumns(columns);
                if (remainingSkip > 0) { remainingSkip--; continue; }

                if (sheet is null || sheet.RowCount == options.MaxRowsPerWorksheet)
                {
                    if (sheet is not null)
                    {
                        if (!options.SplitLargeWorksheets)
                            throw new InvalidOperationException(
                                "The worksheet row limit was exceeded. Enable SplitLargeWorksheets.");
                        sheet.Dispose();
                        sheet = null;
                    }

                    var index = sheetNames.Count + 1;
                    sheet = OpenSheet(zip, index, options, columns);
                    sheetNames.Add(GetSheetName(options.WorksheetName, index));
                }

                // Only one row is materialized at any time. Later unmatched fields
                // are ignored unless the caller supplies an explicit column schema.
                var values = new object?[columns.Count];
                for (var i = 0; i < columns.Count; i++)
                    row.TryGetValue(columns[i], out values[i]);
                sheet.WriteRow(values, header: false);
                written++;
                if (options.MaxRows is long limit && written >= limit) break;
            }

            cancellationToken.ThrowIfCancellationRequested();
            if (sheet is null)
            {
                // An empty input still produces a workbook with one worksheet.
                var index = sheetNames.Count + 1;
                sheet = OpenSheet(zip, index, options, columns ?? []);
                sheetNames.Add(GetSheetName(options.WorksheetName, index));
            }
        }
        finally
        {
            sheet?.Dispose();
        }

        cancellationToken.ThrowIfCancellationRequested();
        WriteWorkbookPart(zip, sheetNames);
        WriteWorkbookRelationships(zip, sheetNames.Count);
        WriteStylesPart(zip);
        WriteRootRelationships(zip);
        WriteContentTypes(zip, sheetNames.Count);
        cancellationToken.ThrowIfCancellationRequested();
    }

    private static string GetSheetName(string baseName, int index)
    {
        var suffix = index == 1 ? "" : $" ({index})";
        if (suffix.Length >= 31)
            throw new InvalidOperationException("Too many Excel worksheets to name safely.");
        return baseName[..Math.Min(baseName.Length, 31 - suffix.Length)] + suffix;
    }

    private static StreamingZipWorksheet OpenSheet(
        ZipArchive zip, int sheetIndex, StreamingExcelOptions options, IReadOnlyList<string> columns)
    {
        var part = zip.CreateEntry($"xl/worksheets/sheet{sheetIndex}.xml", CompressionLevel.Fastest);
        var output = part.Open();
        var writer = CreateXmlWriter(output);
        writer.WriteStartDocument();
        writer.WriteStartElement("worksheet", SpreadsheetNamespace);
        writer.WriteStartElement("sheetData", SpreadsheetNamespace);
        var sheet = new StreamingZipWorksheet(writer, output);
        if (options.IncludeHeaders && columns.Count != 0)
        {
            var headers = new object?[columns.Count];
            for (var i = 0; i < columns.Count; i++)
                headers[i] = options.ColumnHeaders?.TryGetValue(columns[i], out var label) == true
                    ? label : columns[i];
            sheet.WriteRow(headers, header: true);
        }
        return sheet;
    }

    private sealed class StreamingZipWorksheet(XmlWriter writer, Stream part) : IDisposable
    {
        private bool _disposed;
        public int RowCount { get; private set; }

        public void WriteRow(IReadOnlyList<object?> values, bool header)
        {
            writer.WriteStartElement("row", SpreadsheetNamespace);
            writer.WriteAttributeString("r", (++RowCount).ToString(CultureInfo.InvariantCulture));
            for (var i = 0; i < values.Count; i++)
                WriteCell(writer, values[i], i + 1, RowCount, header);
            writer.WriteEndElement();
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try
            {
                writer.WriteEndElement(); // sheetData
                writer.WriteEndElement(); // worksheet
                writer.WriteEndDocument();
                writer.Flush();
            }
            finally
            {
                writer.Dispose();
                part.Dispose();
            }
        }
    }

    private static void WriteCell(XmlWriter writer, object? value, int column, int row, bool header)
    {
        writer.WriteStartElement("c", SpreadsheetNamespace);
        writer.WriteAttributeString("r", ColumnLetter(column) + row.ToString(CultureInfo.InvariantCulture));
        if (header)
            writer.WriteAttributeString("s", "1");

        if (value is null or DBNull)
        {
            writer.WriteEndElement();
            return;
        }

        if (value is bool boolean)
        {
            writer.WriteAttributeString("t", "b");
            writer.WriteElementString("v", SpreadsheetNamespace, boolean ? "1" : "0");
        }
        else if (value is DateTime date && date.Year >= 1900)
        {
            writer.WriteAttributeString("t", "n");
            writer.WriteAttributeString("s", "2");
            writer.WriteElementString("v", SpreadsheetNamespace,
                date.ToOADate().ToString("R", CultureInfo.InvariantCulture));
        }
        else if (value is DateTimeOffset offset && offset.Year >= 1900)
        {
            writer.WriteAttributeString("t", "n");
            writer.WriteAttributeString("s", "2");
            writer.WriteElementString("v", SpreadsheetNamespace,
                offset.DateTime.ToOADate().ToString("R", CultureInfo.InvariantCulture));
        }
        else if (TryGetExcelNumber(value, out var number))
        {
            writer.WriteAttributeString("t", "n");
            writer.WriteElementString("v", SpreadsheetNamespace, number);
        }
        else
        {
            // Inline strings avoid a global shared-strings table and its O(n)
            // retention of unique values. XmlWriter escapes cell text safely.
            writer.WriteAttributeString("t", "inlineStr");
            writer.WriteStartElement("is", SpreadsheetNamespace);
            writer.WriteStartElement("t", SpreadsheetNamespace);
            writer.WriteAttributeString("xml", "space", XmlNamespace, "preserve");
            writer.WriteString(Convert.ToString(value, CultureInfo.InvariantCulture) ?? "");
            writer.WriteEndElement(); // t
            writer.WriteEndElement(); // is
        }

        writer.WriteEndElement(); // c
    }

    private static bool TryGetExcelNumber(object value, out string number)
    {
        // Excel stores a maximum of 15 significant decimal digits reliably;
        // large integer identifiers stay strings instead of being rounded.
        switch (value)
        {
            case byte or sbyte or short or ushort or int or uint or float or double or decimal:
                number = Convert.ToString(value, CultureInfo.InvariantCulture) ?? "0";
                return true;
            case long l when l is >= -999999999999999 and <= 999999999999999:
                number = l.ToString(CultureInfo.InvariantCulture);
                return true;
            case ulong ul when ul <= 999999999999999:
                number = ul.ToString(CultureInfo.InvariantCulture);
                return true;
            default:
                number = "";
                return false;
        }
    }

    private static string ColumnLetter(int column)
    {
        Span<char> buffer = stackalloc char[4];
        var cursor = buffer.Length;
        while (column > 0)
        {
            column--;
            buffer[--cursor] = (char)('A' + column % 26);
            column /= 26;
        }
        return new string(buffer[cursor..]);
    }

    private static XmlWriter CreateXmlWriter(Stream destination) =>
        XmlWriter.Create(destination, new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(false),
            CloseOutput = false,
            Indent = false
        });

    private static void WriteXmlPart(ZipArchive zip, string name, Action<XmlWriter> body)
    {
        var entry = zip.CreateEntry(name, CompressionLevel.Fastest);
        using var part = entry.Open();
        using var writer = CreateXmlWriter(part);
        writer.WriteStartDocument();
        body(writer);
        writer.WriteEndDocument();
    }

    private static void WriteWorkbookPart(ZipArchive zip, IReadOnlyList<string> names) =>
        WriteXmlPart(zip, "xl/workbook.xml", xml =>
        {
            xml.WriteStartElement("workbook", SpreadsheetNamespace);
            xml.WriteAttributeString("xmlns", "r", null, OfficeRelationshipNamespace);
            xml.WriteStartElement("sheets", SpreadsheetNamespace);
            for (var i = 0; i < names.Count; i++)
            {
                xml.WriteStartElement("sheet", SpreadsheetNamespace);
                xml.WriteAttributeString("name", names[i]);
                xml.WriteAttributeString("sheetId", (i + 1).ToString(CultureInfo.InvariantCulture));
                xml.WriteAttributeString("r", "id", OfficeRelationshipNamespace, $"rId{i + 1}");
                xml.WriteEndElement();
            }
            xml.WriteEndElement();
            xml.WriteEndElement();
        });

    private static void WriteWorkbookRelationships(ZipArchive zip, int sheetCount) =>
        WriteXmlPart(zip, "xl/_rels/workbook.xml.rels", xml =>
        {
            xml.WriteStartElement("Relationships", PackageRelationshipNamespace);
            for (var i = 1; i <= sheetCount; i++)
                WriteRelationship(xml, $"rId{i}", OfficeRelationshipBase + "worksheet",
                    $"worksheets/sheet{i}.xml");
            WriteRelationship(xml, $"rId{sheetCount + 1}", OfficeRelationshipBase + "styles", "styles.xml");
            xml.WriteEndElement();
        });

    private static void WriteRootRelationships(ZipArchive zip) =>
        WriteXmlPart(zip, "_rels/.rels", xml =>
        {
            xml.WriteStartElement("Relationships", PackageRelationshipNamespace);
            WriteRelationship(xml, "rId1", OfficeRelationshipBase + "officeDocument", "xl/workbook.xml");
            xml.WriteEndElement();
        });

    private static void WriteRelationship(XmlWriter xml, string id, string type, string target)
    {
        xml.WriteStartElement("Relationship", PackageRelationshipNamespace);
        xml.WriteAttributeString("Id", id);
        xml.WriteAttributeString("Type", type);
        xml.WriteAttributeString("Target", target);
        xml.WriteEndElement();
    }

    private static void WriteContentTypes(ZipArchive zip, int sheetCount) =>
        WriteXmlPart(zip, "[Content_Types].xml", xml =>
        {
            xml.WriteStartElement("Types", ContentTypeNamespace);
            WriteContentType(xml, "Default", "Extension", "rels",
                "application/vnd.openxmlformats-package.relationships+xml");
            WriteContentType(xml, "Default", "Extension", "xml", "application/xml");
            WriteContentType(xml, "Override", "PartName", "/xl/workbook.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
            WriteContentType(xml, "Override", "PartName", "/xl/styles.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
            for (var i = 1; i <= sheetCount; i++)
                WriteContentType(xml, "Override", "PartName", $"/xl/worksheets/sheet{i}.xml",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
            xml.WriteEndElement();
        });

    private static void WriteContentType(
        XmlWriter xml, string element, string identifier, string partName, string contentType)
    {
        xml.WriteStartElement(element, ContentTypeNamespace);
        xml.WriteAttributeString(identifier, partName);
        xml.WriteAttributeString("ContentType", contentType);
        xml.WriteEndElement();
    }

    private static void WriteStylesPart(ZipArchive zip) =>
        WriteXmlPart(zip, "xl/styles.xml", xml =>
        {
            xml.WriteStartElement("styleSheet", SpreadsheetNamespace);

            xml.WriteStartElement("fonts", SpreadsheetNamespace);
            xml.WriteAttributeString("count", "2");
            WriteFont(xml, bold: false);
            WriteFont(xml, bold: true);
            xml.WriteEndElement();

            xml.WriteStartElement("fills", SpreadsheetNamespace);
            xml.WriteAttributeString("count", "2");
            WritePatternFill(xml, "none");
            WritePatternFill(xml, "gray125");
            xml.WriteEndElement();

            xml.WriteStartElement("borders", SpreadsheetNamespace);
            xml.WriteAttributeString("count", "1");
            xml.WriteStartElement("border", SpreadsheetNamespace);
            xml.WriteEndElement();
            xml.WriteEndElement();

            xml.WriteStartElement("cellStyleXfs", SpreadsheetNamespace);
            xml.WriteAttributeString("count", "1");
            WriteXf(xml, fontId: 0, numFmtId: 0, isBold: false, isDate: false);
            xml.WriteEndElement();

            xml.WriteStartElement("cellXfs", SpreadsheetNamespace);
            xml.WriteAttributeString("count", "3");
            WriteXf(xml, fontId: 0, numFmtId: 0, isBold: false, isDate: false);
            WriteXf(xml, fontId: 1, numFmtId: 0, isBold: true, isDate: false);
            WriteXf(xml, fontId: 0, numFmtId: 22, isBold: false, isDate: true);
            xml.WriteEndElement();

            xml.WriteStartElement("cellStyles", SpreadsheetNamespace);
            xml.WriteAttributeString("count", "1");
            xml.WriteStartElement("cellStyle", SpreadsheetNamespace);
            xml.WriteAttributeString("name", "Normal");
            xml.WriteAttributeString("xfId", "0");
            xml.WriteAttributeString("builtinId", "0");
            xml.WriteEndElement();
            xml.WriteEndElement();
            xml.WriteEndElement();
        });

    private static void WriteFont(XmlWriter xml, bool bold)
    {
        xml.WriteStartElement("font", SpreadsheetNamespace);
        if (bold) xml.WriteElementString("b", SpreadsheetNamespace, "");
        xml.WriteStartElement("sz", SpreadsheetNamespace);
        xml.WriteAttributeString("val", "11");
        xml.WriteEndElement();
        xml.WriteStartElement("name", SpreadsheetNamespace);
        xml.WriteAttributeString("val", "Calibri");
        xml.WriteEndElement();
        xml.WriteEndElement();
    }

    private static void WritePatternFill(XmlWriter xml, string pattern)
    {
        xml.WriteStartElement("fill", SpreadsheetNamespace);
        xml.WriteStartElement("patternFill", SpreadsheetNamespace);
        xml.WriteAttributeString("patternType", pattern);
        xml.WriteEndElement();
        xml.WriteEndElement();
    }

    private static void WriteXf(XmlWriter xml, int fontId, int numFmtId, bool isBold, bool isDate)
    {
        xml.WriteStartElement("xf", SpreadsheetNamespace);
        xml.WriteAttributeString("numFmtId", numFmtId.ToString(CultureInfo.InvariantCulture));
        xml.WriteAttributeString("fontId", fontId.ToString(CultureInfo.InvariantCulture));
        xml.WriteAttributeString("fillId", "0");
        xml.WriteAttributeString("borderId", "0");
        xml.WriteAttributeString("xfId", "0");
        if (isBold) xml.WriteAttributeString("applyFont", "1");
        if (isDate) xml.WriteAttributeString("applyNumberFormat", "1");
        xml.WriteEndElement();
    }
}
