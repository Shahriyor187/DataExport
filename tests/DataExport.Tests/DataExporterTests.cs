using ClosedXML.Excel;
using DataExportPackage;
using System.Text.Json;
using Xunit;

namespace DataExport.Tests;

public sealed class DataExporterTests
{
    [Fact]
    public void ToJson_WritesRecordsAndPreservesNull()
    {
        var path = CreateTempFile(".json");

        try
        {
            var exporter = new DataExporter(
            [
                new Dictionary<string, object>
                {
                    ["Name"] = "Amina",
                    ["Age"] = 21,
                    ["Nickname"] = null!
                }
            ]);

            exporter.ToJson(path);

            using var document = JsonDocument.Parse(File.ReadAllText(path));
            var record = document.RootElement[0];

            Assert.Equal("Amina", record.GetProperty("Name").GetString());
            Assert.Equal(21, record.GetProperty("Age").GetInt32());
            Assert.Equal(JsonValueKind.Null, record.GetProperty("Nickname").ValueKind);
        }
        finally
        {
            DeleteIfExists(path);
        }
    }

    [Fact]
    public void ToXml_EncodesInvalidElementNamesAndHandlesNull()
    {
        var path = CreateTempFile(".xml");

        try
        {
            var exporter = new DataExporter(
            [
                new Dictionary<string, object>
                {
                    ["first name"] = "Amina",
                    ["nickname"] = null!
                }
            ]);

            exporter.ToXml(path);

            var xml = File.ReadAllText(path);
            Assert.Contains("<first_x0020_name>Amina</first_x0020_name>", xml);
            Assert.Contains("<nickname></nickname>", xml);
        }
        finally
        {
            DeleteIfExists(path);
        }
    }

    [Fact]
    public void ToCsv_UsesUnionOfKeysAndEscapesFields()
    {
        var path = CreateTempFile(".csv");

        try
        {
            var exporter = new DataExporter(
            [
                new Dictionary<string, object>
                {
                    ["Name"] = "Alice, Smith",
                    ["Note"] = "said \"hello\""
                },
                new Dictionary<string, object>
                {
                    ["Name"] = "Bob",
                    ["Country"] = "Uzbekistan"
                }
            ]);

            exporter.ToCsv(path);

            var csv = File.ReadAllText(path);
            Assert.Contains("Name,Note,Country", csv);
            Assert.Contains("\"Alice, Smith\"", csv);
            Assert.Contains("hello", csv);
            Assert.Contains("Bob,,Uzbekistan", csv);
        }
        finally
        {
            DeleteIfExists(path);
        }
    }

    [Fact]
    public void ToExcel_WritesHeadersAndAlignsRecordsByKey()
    {
        var path = CreateTempFile(".xlsx");

        try
        {
            var exporter = new DataExporter(
            [
                new Dictionary<string, object>
                {
                    ["Name"] = "Amina",
                    ["Age"] = 21
                },
                new Dictionary<string, object>
                {
                    ["Name"] = "Ali",
                    ["Country"] = "Uzbekistan"
                }
            ]);

            exporter.ToExcel(path, includeHeaders: true);

            using var workbook = new XLWorkbook(path);
            var worksheet = workbook.Worksheet("Data");

            Assert.Equal("Name", worksheet.Cell(1, 1).GetString());
            Assert.Equal("Age", worksheet.Cell(1, 2).GetString());
            Assert.Equal("Country", worksheet.Cell(1, 3).GetString());
            Assert.Equal("Amina", worksheet.Cell(2, 1).GetString());
            Assert.Equal("21", worksheet.Cell(2, 2).GetString());
            Assert.Equal(string.Empty, worksheet.Cell(2, 3).GetString());
            Assert.Equal("Uzbekistan", worksheet.Cell(3, 3).GetString());
        }
        finally
        {
            DeleteIfExists(path);
        }
    }

    [Fact]
    public void Constructor_RejectsNullCollection()
    {
        Assert.Throws<ArgumentNullException>(() => new DataExporter(null!));
    }


    [Fact]
    public void ToCsv_HandlesCommasQuotesAndNewlines()
    {
        var path = CreateTempFile(".csv");

        try
        {
            var exporter = new DataExporter(
            [
                new Dictionary<string, object>
                {
                    ["Message"] = "A, \"quoted\" value\nwith a second line"
                }
            ]);

            exporter.ToCsv(path);

            var contents = File.ReadAllText(path);
            Assert.Contains("\"\"quoted\"\"", contents);
            Assert.Contains("with a second line", contents);
        }
        finally
        {
            DeleteIfExists(path);
        }
    }

    [Fact]
    public void ToExcel_DefaultOmitsHeaders()
    {
        var path = CreateTempFile(".xlsx");

        try
        {
            var exporter = new DataExporter(
            [
                new Dictionary<string, object> { ["Name"] = "Amina" }
            ]);
            exporter.ToExcel(path);

            using var workbook = new XLWorkbook(path);
            Assert.Equal("Amina", workbook.Worksheet("Data").Cell(1, 1).GetString());
        }
        finally
        {
            DeleteIfExists(path);
        }
    }

    [Fact]
    public void ToJson_HandlesEmptyData()
    {
        var path = CreateTempFile(".json");

        try
        {
            new DataExporter([]).ToJson(path);
            Assert.Equal("[]", File.ReadAllText(path));
        }
        finally
        {
            DeleteIfExists(path);
        }
    }


    [Fact]
    public void FromObjects_ExportsPropertiesAndTypedExcelCells()
    {
        var path = CreateTempFile(".xlsx");
        try
        {
            var exporter = DataExporter.FromObjects(
                new[] { new { Name = "Amina", Score = 1450, Passed = true } });
            exporter.ToExcel(path, new DataExportOptions
            {
                ExcelIncludeHeaders = true,
                ExcelFreezeHeader = true,
                ExcelAutoFilter = true,
                ExcelAutoFitColumns = false
            });
            using var workbook = new XLWorkbook(path);
            var worksheet = workbook.Worksheet("Data");
            Assert.Equal("Name", worksheet.Cell(1, 1).GetString());
            Assert.Equal(1450, worksheet.Cell(2, 2).GetDouble());
            Assert.True(worksheet.Cell(2, 3).GetBoolean());
        }
        finally { DeleteIfExists(path); }
    }

    [Fact]
    public async Task ToCsvAsync_WritesToOpenStreamAndProtectsFormulaText()
    {
        var exporter = DataExporter.FromObjects(new[]
        {
            new { Name = "Ali", Comment = "=HYPERLINK(\"https://invalid.example\")" }
        });
        using var stream = new MemoryStream();
        await exporter.ToCsvAsync(stream);
        Assert.True(stream.CanWrite);
        var content = System.Text.Encoding.UTF8.GetString(stream.ToArray());
        Assert.Contains("'=HYPERLINK", content);
    }

    [Fact]
    public async Task ExportAsync_WritesJsonToCallerOwnedStream()
    {
        var exporter = DataExporter.FromObjects(new[] { new { Id = 3, Name = "Ali" } });
        using var output = new MemoryStream();
        await exporter.ExportAsync(output, ExportFormat.Json);
        Assert.True(output.CanWrite);
        using var json = JsonDocument.Parse(output.ToArray());
        Assert.Equal(3, json.RootElement[0].GetProperty("Id").GetInt32());
    }

    [Fact]
    public void FromDataTable_ExportsColumnValues()
    {
        var table = new System.Data.DataTable();
        table.Columns.Add("Score", typeof(int));
        table.Rows.Add(1500);
        var path = CreateTempFile(".json");
        try
        {
            DataExporter.FromDataTable(table).ToJson(path);
            using var json = JsonDocument.Parse(File.ReadAllText(path));
            Assert.Equal(1500, json.RootElement[0].GetProperty("Score").GetInt32());
        }
        finally { DeleteIfExists(path); }
    }

    [Fact]
    public async Task WriteCsvAsync_StreamsAsynchronousRecords()
    {
        static async IAsyncEnumerable<CsvRow> Records()
        {
            yield return new CsvRow(1, "Amina");
            await Task.Yield();
            yield return new CsvRow(2, "Ali");
        }

        using var output = new MemoryStream();
        await DataExporter.WriteCsvAsync(Records(), output);
        var csv = System.Text.Encoding.UTF8.GetString(output.ToArray());
        Assert.Contains("Id,Name", csv);
        Assert.Contains("1,Amina", csv);
        Assert.Contains("2,Ali", csv);
    }

    private sealed record CsvRow(int Id, string Name);


    [Fact]
    public void CsvOptions_AllowCustomColumnsDelimiterAndFormulaProtection()
    {
        var path = CreateTempFile(".csv");
        try
        {
            var data = DataExporter.FromObjects(new[]
            {
                new { Name = "Amina", Formula = "=1+2", Score = 1500 }
            });
            data.ToCsv(path, new DataExportOptions
            {
                Columns = ["Formula", "Name"],
                CsvDelimiter = ';'
            });

            var csv = File.ReadAllText(path);
            Assert.Contains("Formula;Name", csv);
            Assert.Contains("'=1+2;Amina", csv);
            Assert.DoesNotContain("Score", csv);
        }
        finally { DeleteIfExists(path); }
    }

    [Fact]
    public void LegacyCsv_DoesNotAlterOriginalCellText()
    {
        var path = CreateTempFile(".csv");
        try
        {
            DataExporter.FromObjects(new[] { new { Formula = "=1+2" } }).ToCsv(path);
            Assert.Contains("=1+2", File.ReadAllText(path));
            Assert.DoesNotContain("'=1+2", File.ReadAllText(path));
        }
        finally { DeleteIfExists(path); }
    }

    [Fact]
    public void ExcelOptions_PreserveDateAndBooleanData()
    {
        var path = CreateTempFile(".xlsx");
        try
        {
            var date = new DateTime(2026, 11, 22);
            var data = DataExporter.FromObjects(new[] { new { Date = date, Active = true } });
            data.ToExcel(path, new DataExportOptions
            {
                WorksheetName = "Summary",
                ExcelIncludeHeaders = true,
                ExcelAutoFitColumns = false
            });
            using var workbook = new XLWorkbook(path);
            Assert.Equal(date, workbook.Worksheet("Summary").Cell(2, 1).GetDateTime());
            Assert.True(workbook.Worksheet("Summary").Cell(2, 2).GetBoolean());
        }
        finally { DeleteIfExists(path); }
    }

    private static string CreateTempFile(string extension) =>
        Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}{extension}");

    private static void DeleteIfExists(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }
}
