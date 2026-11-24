using System.Text;

namespace DataExportPackage;

/// <summary>Output formats supported by <see cref="DataExporter"/>.</summary>
public enum ExportFormat
{
    /// <summary>JavaScript Object Notation.</summary>
    Json,
    /// <summary>Extensible Markup Language.</summary>
    Xml,
    /// <summary>Comma-separated values.</summary>
    Csv,
    /// <summary>Excel Open XML workbook.</summary>
    Excel
}

/// <summary>Optional export settings. The existing file-based methods remain available.</summary>
public sealed class DataExportOptions
{
    /// <summary>Number of data records to skip before exporting. Headers are excluded.</summary>
    public long RowOffset { get; init; }
    /// <summary>Maximum data records to export, or null for unlimited. Zero exports no data.</summary>
    public long? MaxRows { get; init; }

    /// <summary>Explicit column names and order, or null to use all discovered columns.</summary>
    public IReadOnlyList<string>? Columns { get; set; }

    /// <summary>Field separator for CSV; comma by default.</summary>
    public char CsvDelimiter { get; set; } = ',';
    /// <summary>Include the column names as the first CSV row.</summary>
    public bool CsvIncludeHeaders { get; set; } = true;

    /// <summary>Text encoding used for CSV. Null uses UTF-8 without a BOM.</summary>
    public Encoding? CsvEncoding { get; set; }

    /// <summary>Prefix text that Excel could interpret as a formula with an apostrophe.</summary>
    public bool CsvProtectFormulas { get; set; } = true;

    /// <summary>Name of the exported Excel worksheet.</summary>
    public string WorksheetName { get; set; } = "Data";
    /// <summary>Write column names in the first Excel row.</summary>
    public bool ExcelIncludeHeaders { get; set; }
    /// <summary>Use bold font for Excel headers.</summary>
    public bool ExcelBoldHeaders { get; set; } = true;
    /// <summary>Freeze the first row when headers are included.</summary>
    public bool ExcelFreezeHeader { get; set; }
    /// <summary>Enable an Excel filter on the header range.</summary>
    public bool ExcelAutoFilter { get; set; }
    /// <summary>Automatically size worksheet columns to their contents.</summary>
    public bool ExcelAutoFitColumns { get; set; } = true;
}
