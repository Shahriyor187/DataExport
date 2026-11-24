using ClosedXML.Excel;
using DataExportPackage;
using System.Data;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Xunit;

namespace DataExport.Tests;

public sealed class RowWindowTests
{
    private sealed record Row(int Id);

    [Theory]
    [InlineData(ExportFormat.Json)]
    [InlineData(ExportFormat.Xml)]
    [InlineData(ExportFormat.Csv)]
    [InlineData(ExportFormat.Excel)]
    public async Task BufferedWindowPreservesOrderAndHeaders(ExportFormat format)
    {
        var exporter = DataExporter.FromObjects(Enumerable.Range(1, 8).Select(i => new Row(i)));
        using var output = new MemoryStream();
        await exporter.ExportAsync(output, format, new DataExportOptions
        {
            RowOffset = 2, MaxRows = 3, ExcelIncludeHeaders = true, ExcelAutoFitColumns = false
        });
        Assert.True(output.CanWrite);
        output.Position = 0;
        switch (format)
        {
            case ExportFormat.Json:
                using (var json = JsonDocument.Parse(output))
                    Assert.Equal(new[] { 3, 4, 5 }, json.RootElement.EnumerateArray().Select(x => x.GetProperty("Id").GetInt32()));
                break;
            case ExportFormat.Xml:
                Assert.Equal(new[] { "3", "4", "5" }, XDocument.Load(output).Root!.Elements("entry").Select(x => x.Element("Id")!.Value));
                break;
            case ExportFormat.Csv:
                using (var reader = new StreamReader(output, leaveOpen: true))
                    Assert.Equal(new[] { "Id", "3", "4", "5" }, (await reader.ReadToEndAsync()).Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries));
                break;
            case ExportFormat.Excel:
                using (var book = new XLWorkbook(output))
                {
                    var sheet = book.Worksheet(1);
                    Assert.Equal("Id", sheet.Cell(1, 1).GetString());
                    Assert.Equal(4, sheet.LastRowUsed()!.RowNumber());
                    Assert.Equal(new[] { 3, 4, 5 }, Enumerable.Range(2, 3).Select(i => (int)sheet.Cell(i, 1).GetDouble()));
                }
                break;
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task StreamingStopsWithoutReadingExtraRecord(bool excel)
    {
        int reads = 0;
        async IAsyncEnumerable<Row> Rows()
        {
            for (int i = 1; i <= 5; i++)
            {
                reads++;
                yield return new Row(i);
                await Task.CompletedTask;
            }
            throw new InvalidOperationException("Read past requested window");
        }
        using var output = new MemoryStream();
        if (excel)
            await DataExporter.WriteExcelStreamingAsync(Rows(), output, new StreamingExcelOptions { RowOffset = 2, MaxRows = 3, MaxRowsPerWorksheet = 3 });
        else
            await DataExporter.WriteCsvAsync(Rows(), output, new DataExportOptions { RowOffset = 2, MaxRows = 3 });
        Assert.Equal(5, reads);
        output.Position = 0;
        if (excel)
        {
            using var book = new XLWorkbook(output);
            Assert.Equal(2, book.Worksheets.Count);
            Assert.Equal(3d, book.Worksheet(1).Cell(2, 1).GetDouble());
            Assert.Equal("Id", book.Worksheet(2).Cell(1, 1).GetString());
            Assert.Equal(5d, book.Worksheet(2).Cell(2, 1).GetDouble());
        }
        else
            Assert.Equal(new[] { "Id", "3", "4", "5" }, Encoding.UTF8.GetString(output.ToArray()).Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries));
    }

    [Fact]
    public async Task DatabaseReaderDoesNotAdvancePastLimit()
    {
        using var table = new DataTable();
        table.Columns.Add("Id", typeof(int));
        for (int i = 1; i <= 8; i++) table.Rows.Add(i);
        using var reader = table.CreateDataReader();
        using var output = new MemoryStream();
        await DataExporter.WriteExcelStreamingAsync(reader, output, new StreamingExcelOptions { RowOffset = 2, MaxRows = 3 });
        Assert.False(reader.IsClosed);
        Assert.True(reader.Read());
        Assert.Equal(6, reader.GetInt32(0));
        output.Position = 0;
        using var book = new XLWorkbook(output);
        Assert.Equal(5d, book.Worksheet(1).Cell(4, 1).GetDouble());
    }

    [Theory]
    [InlineData(0, 0)]
    [InlineData(long.MaxValue, 3)]
    public async Task EmptyWindowKeepsKnownHeaders(long offset, long limit)
    {
        var exporter = DataExporter.FromObjects(new[] { new Row(1) });
        using var output = new MemoryStream();
        await exporter.ExportAsync(output, ExportFormat.Csv, new DataExportOptions { RowOffset = offset, MaxRows = limit });
        Assert.Equal("Id", Encoding.UTF8.GetString(output.ToArray()).Trim());
    }

    [Theory]
    [InlineData(-1, 1)]
    [InlineData(0, -1)]
    public async Task InvalidWindowDoesNotWrite(long offset, long limit)
    {
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => DataExporter.FromObjects(new[] { new Row(1) }).ExportAsync(output, ExportFormat.Json, new DataExportOptions { RowOffset = offset, MaxRows = limit }));
        Assert.Equal(0, output.Length);
    }
}
