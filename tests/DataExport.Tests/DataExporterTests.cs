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

            exporter.ToExcel(path);

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
            Assert.Contains("quoted", contents);
            Assert.Contains("with a second line", contents);
        }
        finally
        {
            DeleteIfExists(path);
        }
    }

    [Fact]
    public void ToExcel_CanOmitHeaders()
    {
        var path = CreateTempFile(".xlsx");

        try
        {
            var exporter = new DataExporter(
            [
                new Dictionary<string, object> { ["Name"] = "Amina" }
            ]);
            exporter.ToExcel(path, includeHeaders: false);

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
