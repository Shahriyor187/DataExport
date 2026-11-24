using System.Data;
using DataExportPackage;
using Xunit;

namespace DataExport.Tests;

public sealed class CancellationTests
{
    [Theory]
    [InlineData(ExportFormat.Json)]
    [InlineData(ExportFormat.Xml)]
    [InlineData(ExportFormat.Csv)]
    [InlineData(ExportFormat.Excel)]
    public async Task PreCancelledExportDoesNotWrite(ExportFormat format)
    {
        using var source = new CancellationTokenSource();
        source.Cancel();
        using var stream = new MemoryStream();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            DataExporter.FromObjects(new[] { new { Id = 1 } }).ExportAsync(stream, format, cancellationToken: source.Token));
        Assert.True(stream.CanWrite);
        Assert.Equal(0, stream.Length);
    }

    [Fact]
    public async Task PreCancelledStreamingDoesNotEnumerateOrWrite()
    {
        using var source = new CancellationTokenSource();
        source.Cancel();
        var visited = false;
        async IAsyncEnumerable<int> Rows()
        {
            visited = true;
            yield return 1;
            await Task.CompletedTask;
        }
        using var stream = new MemoryStream();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            DataExporter.WriteCsvAsync(Rows(), stream, cancellationToken: source.Token));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            DataExporter.WriteExcelStreamingAsync(Rows(), stream, cancellationToken: source.Token));
        using var table = new DataTable();
        table.Columns.Add("Id", typeof(int));
        using var reader = table.CreateDataReader();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            DataExporter.WriteExcelStreamingAsync(reader, stream, cancellationToken: source.Token));
        Assert.False(visited);
        Assert.False(reader.IsClosed);
        Assert.True(stream.CanWrite);
        Assert.Equal(0, stream.Length);
    }

    [Fact]
    public async Task StreamingCsvChecksCancellationEvenIfSourceIgnoresToken()
    {
        using var source = new CancellationTokenSource();
        using var stream = new MemoryStream();
        var visited = 0;
        async IAsyncEnumerable<Row> Rows()
        {
            for (var i = 0; i < 10; i++)
            {
                visited++;
                if (i == 2) source.Cancel();
                yield return new Row(i);
                await Task.CompletedTask;
            }
        }
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            DataExporter.WriteCsvAsync(Rows(), stream, cancellationToken: source.Token));
        Assert.Equal(3, visited);
        Assert.True(stream.CanWrite);
    }

    [Theory]
    [InlineData(ExportFormat.Xml)]
    [InlineData(ExportFormat.Excel)]
    public async Task CancellationDuringCellFormattingStopsLaterRows(ExportFormat format)
    {
        using var source = new CancellationTokenSource();
        using var stream = new MemoryStream();
        var visits = 0;
        var rows = Enumerable.Range(0, 8).Select(i => new Dictionary<string, object>
        {
            ["Value"] = new CallbackValue(() =>
            {
                visits++;
                if (i == 1) source.Cancel();
                return "value";
            })
        }).ToList();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            new DataExporter(rows).ExportAsync(stream, format,
                new DataExportOptions { ExcelAutoFitColumns = false }, source.Token));
        Assert.Equal(2, visits);
        Assert.True(stream.CanWrite);
    }

    private sealed record Row(int Id);
    private sealed class CallbackValue(Func<string> callback)
    {
        public override string ToString() => callback();
    }
}
