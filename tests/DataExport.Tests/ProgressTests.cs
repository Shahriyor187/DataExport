using System.Data;
using DataExportPackage;
using Xunit;

namespace DataExport.Tests;

public sealed class ProgressTests
{
    private sealed record Row(int Id);
    private sealed class Observer(Action<ExportProgress>? action = null) : IProgress<ExportProgress>
    {
        public List<ExportProgress> Reports { get; } = new();
        public void Report(ExportProgress value) { Reports.Add(value); action?.Invoke(value); }
    }

    private static async IAsyncEnumerable<Row> Rows(int count)
    {
        for (int i = 1; i <= count; i++) { yield return new Row(i); await Task.Yield(); }
    }

    [Theory]
    [InlineData(ExportFormat.Json)]
    [InlineData(ExportFormat.Xml)]
    [InlineData(ExportFormat.Csv)]
    [InlineData(ExportFormat.Tsv)]
    [InlineData(ExportFormat.Excel)]
    public async Task BufferedProgressCountsSelectedRows(ExportFormat format)
    {
        var observer = new Observer();
        using var output = new MemoryStream();
        await DataExporter.FromObjects(Enumerable.Range(1, 8).Select(i => new Row(i)))
            .ExportAsync(output, format, new DataExportOptions
            {
                Progress = observer, ProgressIntervalRows = 2, RowOffset = 2, MaxRows = 5,
                TotalRows = 999, ExcelAutoFitColumns = false
            });
        AssertReports(observer, new long[] { 2, 4, 5 }, 5);
        Assert.True(output.CanWrite);
    }

    [Theory]
    [InlineData(ExportFormat.Json)]
    [InlineData(ExportFormat.Xml)]
    [InlineData(ExportFormat.Csv)]
    [InlineData(ExportFormat.Tsv)]
    [InlineData(ExportFormat.Excel)]
    public void SynchronousFileExportsReportProgress(ExportFormat format)
    {
        var dir = Directory.CreateTempSubdirectory();
        try
        {
            var observer = new Observer();
            DataExporter.Export(Enumerable.Range(1, 5).Select(i => new Row(i)),
                Path.Combine(dir.FullName, "output"), format, new DataExportOptions
                {
                    Progress = observer, ProgressIntervalRows = 2, ExcelAutoFitColumns = false
                });
            AssertReports(observer, new long[] { 2, 4, 5 }, 5);
        }
        finally { dir.Delete(true); }
    }

    [Theory]
    [InlineData("csv")]
    [InlineData("tsv")]
    [InlineData("xlsx")]
    [InlineData("db-csv")]
    [InlineData("db-tsv")]
    [InlineData("db-xlsx")]
    public async Task StreamingProgressUsesUnknownTotalAndGlobalRowCount(string format)
    {
        var observer = new Observer();
        using var output = new MemoryStream();
        using var table = new DataTable();
        table.Columns.Add("Id", typeof(int));
        for (var i = 1; i <= 8; i++) table.Rows.Add(i);
        using var reader = table.CreateDataReader();
        var options = new DataExportOptions
        {
            Progress = observer, ProgressIntervalRows = 2, RowOffset = 2, MaxRows = 5
        };
        var excel = new StreamingExcelOptions
        {
            Progress = observer, ProgressIntervalRows = 2, RowOffset = 2, MaxRows = 5,
            MaxRowsPerWorksheet = 3
        };
        switch (format)
        {
            case "csv": await DataExporter.WriteCsvAsync(Rows(8), output, options); break;
            case "tsv": await DataExporter.WriteTsvAsync(Rows(8), output, options); break;
            case "xlsx": await DataExporter.WriteExcelStreamingAsync(Rows(8), output, excel); break;
            case "db-csv": await DataExporter.WriteCsvAsync(reader, output, options); break;
            case "db-tsv": await DataExporter.WriteTsvAsync(reader, output, options); break;
            default: await DataExporter.WriteExcelStreamingAsync(reader, output, excel); break;
        }
        AssertReports(observer, new long[] { 2, 4, 5 }, null);
        Assert.True(output.CanWrite);
        Assert.False(reader.IsClosed);
        if (format.StartsWith("db-"))
        {
            Assert.True(reader.Read());
            Assert.Equal(8, reader.GetInt32(0));
        }
    }

    [Fact]
    public async Task IncorrectEstimateBecomesUnknownAndExactIntervalIsNotDuplicated()
    {
        var observer = new Observer();
        using var output = new MemoryStream();
        await DataExporter.WriteCsvAsync(Rows(4), output, new DataExportOptions
        { Progress = observer, ProgressIntervalRows = 2, TotalRows = 2 });
        Assert.Equal(new long[] { 2, 4 }, observer.Reports.Select(p => p.RowsWritten));
        Assert.Equal(2L, observer.Reports[0].TotalRows);
        Assert.Null(observer.Reports[1].TotalRows);
    }

    [Fact]
    public async Task ZeroRowsReportZeroWithoutEnumeratingSource()
    {
        var observer = new Observer();
        using var output = new MemoryStream();
        await DataExporter.WriteCsvAsync(Rows(5), output, new DataExportOptions
        { Progress = observer, MaxRows = 0, Columns = new[] { "Id" } });
        AssertReports(observer, new long[] { 0 }, 0);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CallbackCancellationStopsBeforeNextRecord(bool excel)
    {
        using var cancel = new CancellationTokenSource();
        var observer = new Observer(_ => cancel.Cancel());
        using var output = new MemoryStream();
        var reads = 0;
        async IAsyncEnumerable<Row> Source()
        {
            for (int i = 1; i <= 8; i++) { reads++; yield return new Row(i); await Task.Yield(); }
        }
        var task = excel
            ? DataExporter.WriteExcelStreamingAsync(Source(), output, new StreamingExcelOptions
            { Progress = observer, ProgressIntervalRows = 2 }, cancel.Token)
            : DataExporter.WriteCsvAsync(Source(), output, new DataExportOptions
            { Progress = observer, ProgressIntervalRows = 2 }, cancel.Token);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => task);
        Assert.Equal(2, reads);
        AssertReports(observer, new long[] { 2 }, null);
        Assert.True(output.CanWrite);
    }

    [Fact]
    public async Task CallbackFailurePropagatesAndLeavesStreamOpen()
    {
        var observer = new Observer(_ => throw new InvalidOperationException("callback failure"));
        using var output = new MemoryStream();
        var error = await Assert.ThrowsAsync<InvalidOperationException>(() =>
            DataExporter.WriteTsvAsync(Rows(5), output, new DataExportOptions
            { Progress = observer, ProgressIntervalRows = 2 }));
        Assert.Equal("callback failure", error.Message);
        AssertReports(observer, new long[] { 2 }, null);
        Assert.True(output.CanWrite);
    }

    [Fact]
    public async Task ConcurrentExportsHaveIndependentCounts()
    {
        var a = new Observer();
        var b = new Observer();
        using var first = new MemoryStream();
        using var second = new MemoryStream();
        await Task.WhenAll(
            DataExporter.WriteCsvAsync(Rows(5), first, new DataExportOptions { Progress = a, ProgressIntervalRows = 2 }),
            DataExporter.WriteCsvAsync(Rows(7), second, new DataExportOptions { Progress = b, ProgressIntervalRows = 3 }));
        AssertReports(a, new long[] { 2, 4, 5 }, null);
        AssertReports(b, new long[] { 3, 6, 7 }, null);
    }

    [Theory]
    [InlineData(0, null)]
    [InlineData(-1, null)]
    [InlineData(1, -1)]
    public async Task InvalidConfigurationRejectedBeforeOutput(int interval, int? total)
    {
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => DataExporter.WriteCsvAsync(Rows(5), output,
            new DataExportOptions { ProgressIntervalRows = interval, TotalRows = total }));
        Assert.Equal(0, output.Length);
    }

    private static void AssertReports(Observer observer, long[] rows, long? total)
    {
        Assert.Equal(rows, observer.Reports.Select(p => p.RowsWritten));
        Assert.All(observer.Reports, p => { Assert.Equal(total, p.TotalRows); Assert.True(p.Elapsed >= TimeSpan.Zero); });
        for (int i = 1; i < observer.Reports.Count; i++)
            Assert.True(observer.Reports[i].Elapsed >= observer.Reports[i - 1].Elapsed);
    }
}
