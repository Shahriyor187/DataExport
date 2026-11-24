using DataExportPackage;
using System.Text;
using Xunit;

namespace DataExport.Tests;

public sealed class TsvTests
{
    private sealed record Row(int Id, string Text);

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task TsvUsesTabsAndEscapesFieldsWithoutMutatingOptions(bool generic)
    {
        var exporter = DataExporter.FromObjects(new[]
        {
            new Row(1, "skip"),
            new Row(2, "a\tb\"c\nnext"),
            new Row(3, "=1+1"),
            new Row(4, "exclude")
        });
        var options = new DataExportOptions { Columns = new[] { "Id", "Text" }, CsvDelimiter = ';', RowOffset = 1, MaxRows = 2 };
        using var output = new MemoryStream();
        if (generic)
            await exporter.ExportAsync(output, ExportFormat.Tsv, options);
        else
            await exporter.ToTsvAsync(output, options);
        var nl = Environment.NewLine;
        Assert.Equal("Id\tText" + nl + "2\t\"a\tb\"\"c\nnext\"" + nl + "3\t'=1+1" + nl, Encoding.UTF8.GetString(output.ToArray()));
        Assert.Equal(';', options.CsvDelimiter);
        Assert.True(output.CanWrite);
    }

    [Fact]
    public void FileAndGenericExportAgreeAndPreserveEnumValues()
    {
        var dir = Directory.CreateTempSubdirectory();
        try
        {
            var rows = new[] { new Row(1, "a,b"), new Row(2, "é") };
            var a = Path.Combine(dir.FullName, "a.tsv");
            var b = Path.Combine(dir.FullName, "b.tsv");
            DataExporter.FromObjects(rows).ToTsv(a);
            DataExporter.Export(rows, b, ExportFormat.Tsv);
            Assert.Equal(File.ReadAllBytes(a), File.ReadAllBytes(b));
            Assert.Contains("1\ta,b", File.ReadAllText(a));
            Assert.Equal(0, (int)ExportFormat.Json);
            Assert.Equal(1, (int)ExportFormat.Xml);
            Assert.Equal(2, (int)ExportFormat.Csv);
            Assert.Equal(3, (int)ExportFormat.Excel);
            Assert.Equal(4, (int)ExportFormat.Tsv);
        }
        finally { dir.Delete(true); }
    }

    [Fact]
    public async Task StreamingHonorsWindowAndDoesNotReadExtraRecord()
    {
        var visited = 0;
        async IAsyncEnumerable<Row> Rows()
        {
            for (int i = 1; i <= 3; i++)
            {
                visited++;
                yield return new Row(i, "value");
                await Task.CompletedTask;
            }
            throw new InvalidOperationException("Unexpected extra read");
        }
        using var output = new MemoryStream();
        await DataExporter.WriteTsvAsync(Rows(), output, new DataExportOptions { RowOffset = 1, MaxRows = 2, CsvIncludeHeaders = false });
        Assert.Equal(3, visited);
        Assert.Equal("2\tvalue" + Environment.NewLine + "3\tvalue" + Environment.NewLine, Encoding.UTF8.GetString(output.ToArray()));
        Assert.True(output.CanWrite);
    }

    [Fact]
    public async Task StreamingZeroRowsUsesKnownHeadersWithoutEnumerating()
    {
        var visited = false;
        async IAsyncEnumerable<Row> Rows()
        {
            visited = true;
            yield return new Row(1, "unexpected");
            await Task.CompletedTask;
        }
        using var output = new MemoryStream();
        await DataExporter.WriteTsvAsync(Rows(), output, new DataExportOptions { MaxRows = 0, Columns = new[] { "Id", "Text" } });
        Assert.False(visited);
        Assert.Equal("Id\tText" + Environment.NewLine, Encoding.UTF8.GetString(output.ToArray()));
    }

    [Fact]
    public async Task CancellationAndEncodingArePreserved()
    {
        var exporter = DataExporter.FromObjects(new[] { new Row(1, "é") });
        using var output = new MemoryStream();
        using var cancel = new CancellationTokenSource();
        cancel.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => exporter.ToTsvAsync(output, cancellationToken: cancel.Token));
        Assert.True(output.CanWrite);
        Assert.Equal(0, output.Length);
        await exporter.ToTsvAsync(output, new DataExportOptions { CsvEncoding = Encoding.Unicode, CsvIncludeHeaders = false });
        Assert.StartsWith("1\té", Encoding.Unicode.GetString(output.ToArray()).TrimStart('\uFEFF'));
    }

    [Fact]
    public async Task ExplicitFormulaProtectionOptOutIsRespected()
    {
        using var output = new MemoryStream();
        await DataExporter.FromObjects(new[] { new Row(1, "=1+1") }).ToTsvAsync(output, new DataExportOptions { CsvProtectFormulas = false });
        Assert.Contains("1\t=1+1", Encoding.UTF8.GetString(output.ToArray()));
    }
}
