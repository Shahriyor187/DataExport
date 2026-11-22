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

    [Fact]
    public async Task WritesToWriteOnlyFileStream()
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        try
        {
            await using (var output = new FileStream(path, FileMode.CreateNew,
                FileAccess.Write, FileShare.None))
                await DataExporter.WriteExcelStreamingAsync(Generate(100), output);
            using var workbook = new XLWorkbook(path);
            Assert.Equal(101, workbook.Worksheet(1).LastRowUsed()!.RowNumber());
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public async Task WritesValidWorkbookToNonSeekableOutput()
    {
        using var storage = new MemoryStream();
        using (var output = new NonSeekableWriteStream(storage))
            await DataExporter.WriteExcelStreamingAsync(Generate(25), output);

        Assert.True(storage.CanWrite);
        storage.Position = 0;
        using var workbook = new XLWorkbook(storage);
        Assert.Equal(26, workbook.Worksheet(1).LastRowUsed()!.RowNumber());
        Assert.Equal("Student 24", workbook.Worksheet(1).Cell("B26").GetString());
    }

    private sealed class NonSeekableWriteStream(Stream inner) : Stream
    {
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => throw new NotSupportedException();
        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }
        public override void Flush() => inner.Flush();
        public override Task FlushAsync(CancellationToken cancellationToken) =>
            inner.FlushAsync(cancellationToken);
        public override int Read(byte[] buffer, int offset, int count) =>
            throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) =>
            throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) =>
            inner.Write(buffer, offset, count);
        public override void Write(ReadOnlySpan<byte> buffer) => inner.Write(buffer);
        public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer,
            CancellationToken cancellationToken = default) =>
            inner.WriteAsync(buffer, cancellationToken);
    }

    private static async IAsyncEnumerable<Row> Generate(int count)
    {
        for (var i = 0; i < count; i++)
            yield return new Row(i, "Student " + i, i % 2 == 0);
        await Task.CompletedTask;
    }
}
