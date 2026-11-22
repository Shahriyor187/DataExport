namespace DataExportPackage;

/// <summary>Output formats supported by <see cref="DataExporter"/>.</summary>
public enum ExportFormat
{
    Json,
    Xml,
    Csv,
    Excel
}

/// <summary>Optional export settings. The existing file-based methods remain available.</summary>
public sealed class DataExportOptions
{
    /// <summary>Explicit column names and order, or null to use all discovered columns.</summary>
    public IReadOnlyList<string>? Columns { get; set; }

    public char CsvDelimiter { get; set; } = ',';
    public bool CsvIncludeHeaders { get; set; } = true;

    /// <summary>Prefix text that Excel could interpret as a formula with an apostrophe.</summary>
    public bool CsvProtectFormulas { get; set; } = true;

    public string WorksheetName { get; set; } = "Data";
    public bool ExcelIncludeHeaders { get; set; }
    public bool ExcelBoldHeaders { get; set; } = true;
    public bool ExcelFreezeHeader { get; set; }
    public bool ExcelAutoFilter { get; set; }
    public bool ExcelAutoFitColumns { get; set; } = true;
}
