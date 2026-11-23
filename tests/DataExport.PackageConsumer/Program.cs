using ClosedXML.Excel;
using DataExportPackage;
using System.Text.Json;
using System.Xml.Linq;

var root = Path.Combine(Path.GetTempPath(), "DataExport-consumer-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(root);
try
{
    var rows = new[]
    {
        new { Name = "Amina", Score = 1450, Active = true },
        new { Name = "Ali", Score = 1400, Active = false }
    };
    var exporter = DataExporter.FromObjects(rows);

    var json = Path.Combine(root, "records.json");
    var xml = Path.Combine(root, "records.xml");
    var csv = Path.Combine(root, "records.csv");
    var xlsx = Path.Combine(root, "records.xlsx");
    exporter.ToJson(json);
    exporter.ToXml(xml);
    exporter.ToCsv(csv, new DataExportOptions { CsvProtectFormulas = true });
    exporter.ToExcel(xlsx, new DataExportOptions { ExcelIncludeHeaders = true });

    using (var document = JsonDocument.Parse(File.ReadAllText(json)))
    {
        if (document.RootElement.GetArrayLength() != 2 ||
            document.RootElement[0].GetProperty("Score").GetInt32() != 1450)
            throw new Exception("JSON content mismatch.");
    }
    var xmlDoc = XDocument.Load(xml);
    if (xmlDoc.Root?.Name.LocalName != "data" ||
        xmlDoc.Root.Elements("entry").Count() != 2)
        throw new Exception("XML content mismatch.");
    if (!File.ReadAllText(csv).Contains("Amina,1450,True", StringComparison.Ordinal))
        throw new Exception("CSV content mismatch.");
    using (var book = new XLWorkbook(xlsx))
    {
        var sheet = book.Worksheet(1);
        if (sheet.Cell("A1").GetString() != "Name" ||
            sheet.Cell("A2").GetString() != "Amina" ||
            sheet.Cell("B3").GetDouble() != 1400)
            throw new Exception("Excel content mismatch.");
    }

    var streamedPath = Path.Combine(root, "streamed.xlsx");
    await using (var stream = File.Create(streamedPath))
        await DataExporter.WriteExcelStreamingAsync(Generate(), stream,
            new StreamingExcelOptions { Columns = ["Name", "Score", "Active"] });
    using (var book = new XLWorkbook(streamedPath))
    {
        var sheet = book.Worksheet(1);
        if (sheet.LastRowUsed()!.RowNumber() != 3 ||
            sheet.Cell("A3").GetString() != "Ali" ||
            sheet.Cell("B2").GetDouble() != 1450)
            throw new Exception("Streaming XLSX content mismatch.");
    }

    using (var report = new ExcelReportBuilder())
    {
        report.AddSheet("Results", rows);
        report.SaveAs(Path.Combine(root, "report.xlsx"));
    }
    using (var book = new XLWorkbook(Path.Combine(root, "report.xlsx")))
    {
        if (book.Worksheet("Results").Cell("A2").GetString() != "Amina")
            throw new Exception("Report builder content mismatch.");
    }

    Console.WriteLine("Packed DataExport consumer OK: " +
        System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription);
}
finally
{
    if (Directory.Exists(root))
        Directory.Delete(root, recursive: true);
}

static async IAsyncEnumerable<ConsumerRow> Generate()
{
    yield return new ConsumerRow("Amina", 1450, true);
    yield return new ConsumerRow("Ali", 1400, false);
    await Task.CompletedTask;
}

public sealed record ConsumerRow(string Name, int Score, bool Active);
