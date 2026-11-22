using ClosedXML.Excel;
using DataExportPackage;
using System.Runtime.CompilerServices;
using Xunit;

namespace DataExport.Tests;

public sealed class StreamingExcelReliabilityTests
{
    private sealed record Row(int Id, string Name, bool Active);

    [Fact]
    public async Task TenThousandRowsRemainReadableAndCorrect()
    {
        using var output = new MemoryStream();
        await DataExporter.WriteExcelStreamingAsync(Generate(10_000), output,
            new StreamingExcelOptions { Columns = ["Id", "Name", "Active"] });

        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        var worksheet = workbook.Worksheet(1);
        Assert.Equal(10_001, worksheet.LastRowUsed()!.RowNumber());
        Assert.Equal(0, worksheet.Cell("A2").GetDouble());
        Assert.Equal(9_999, worksheet.Cell("A10001").GetDouble());
        Assert.Equal("Student 9999", worksheet.Cell("B10001").GetString());
        Assert.False(worksheet.Cell("C10001").GetBoolean());
    }

    [Fact]
    public async Task CancellationStopsSourceBeforeFullEnumeration()
    {
        using var cancel = new CancellationTokenSource();
        using var output = new MemoryStream();
        var visits = 0;

        async IAsyncEnumerable<Row> Source([EnumeratorCancellation] CancellationToken ct = default)
        {
            for (var i = 0; i < 1_000; i++)
            {
                ct.ThrowIfCancellationRequested();
                visits++;
                if (i == 20) cancel.Cancel();
                yield return new Row(i, "Student " + i, i % 2 == 0);
                await Task.Yield();
            }
        }

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            DataExporter.WriteExcelStreamingAsync(Source(cancel.Token), output,
                cancellationToken: cancel.Token));
        Assert.InRange(visits, 20, 22);
        Assert.True(output.CanWrite);
    }

    [Fact]
    public async Task EmptySourceProducesValidHeaderOnlyWorksheet()
    {
        using var output = new MemoryStream();
        await DataExporter.WriteExcelStreamingAsync(Generate(0), output,
            new StreamingExcelOptions { Columns = ["Id", "Name"] });
        output.Position = 0;
        using var workbook = new XLWorkbook(output);
        Assert.Equal("Id", workbook.Worksheet(1).Cell("A1").GetString());
        Assert.Equal("Name", workbook.Worksheet(1).Cell("B1").GetString());
    }

    private static async IAsyncEnumerable<Row> Generate(int count)
    {
        for (var i = 0; i < count; i++)
            yield return new Row(i, "Student " + i, i % 2 == 0);
        await Task.CompletedTask;
    }
}
