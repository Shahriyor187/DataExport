using ClosedXML.Excel;
using DataExportPackage;
using System.Data;
using Xunit;

namespace DataExport.Tests;

public sealed class StreamingExcelTests
{
    private sealed record Student(string Name, int Score, bool Passed);

    [Fact]
    public async Task StreamsAsyncObjectsWithNativeCellTypes()
    {
        using var output = new MemoryStream();
        await DataExporter.WriteExcelStreamingAsync(GetStudents(), output, new StreamingExcelOptions
        {
            Columns = ["Name", "Score", "Passed"],
            ColumnHeaders = new Dictionary<string, string> { ["Score"] = "SAT Score" }
        });

        Assert.True(output.CanWrite); // stream ownership belongs to caller
        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        var sheet = workbook.Worksheets.Single();
        Assert.Equal("SAT Score", sheet.Cell("B1").GetString());
        Assert.Equal("Amina", sheet.Cell("A2").GetString());
        Assert.Equal(1500, sheet.Cell("B2").GetDouble());
        Assert.True(sheet.Cell("C2").GetBoolean());
        Assert.Equal(1420, sheet.Cell("B3").GetDouble());
    }

    [Fact]
    public async Task RollsOverWorksheetsWithRepeatedHeaders()
    {
        using var output = new MemoryStream();
        await DataExporter.WriteExcelStreamingAsync(GetStudents(), output, new StreamingExcelOptions
        {
            Columns = ["Name", "Score"],
            MaxRowsPerWorksheet = 2,
            WorksheetName = "Exam results"
        });

        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        Assert.Equal(2, workbook.Worksheets.Count);
        Assert.Equal("Exam results", workbook.Worksheets.First().Name);
        Assert.Equal("Exam results (2)", workbook.Worksheets.Last().Name);
        Assert.Equal("Name", workbook.Worksheets.Last().Cell("A1").GetString());
        Assert.Equal("Ali", workbook.Worksheets.Last().Cell("A2").GetString());
    }

    [Fact]
    public async Task StreamsDataReaderWithoutClosingReader()
    {
        var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Score", typeof(int));
        table.Rows.Add("Amina", 1500);
        using var reader = table.CreateDataReader();
        using var output = new MemoryStream();
        await DataExporter.WriteExcelStreamingAsync(reader, output);
        Assert.False(reader.IsClosed);
        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        Assert.Equal(1500, workbook.Worksheet(1).Cell("B2").GetDouble());
    }

    [Fact]
    public async Task RejectsRolloverWhenDisabled()
    {
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<InvalidOperationException>(() =>
            DataExporter.WriteExcelStreamingAsync(GetStudents(), output, new StreamingExcelOptions
            {
                MaxRowsPerWorksheet = 2,
                SplitLargeWorksheets = false
            }));
    }

    private static async IAsyncEnumerable<Student> GetStudents()
    {
        yield return new Student("Amina", 1500, true);
        await Task.Yield();
        yield return new Student("Ali", 1420, false);
    }
}
