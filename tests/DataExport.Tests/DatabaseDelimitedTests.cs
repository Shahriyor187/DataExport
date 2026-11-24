using System.Collections;
using System.Data;
using System.Data.Common;
using System.Text;
using DataExportPackage;
using Xunit;

namespace DataExport.Tests;

public sealed class DatabaseDelimitedTests
{
    private static DataTable Table()
    {
        var table = new DataTable();
        table.Columns.Add("Id", typeof(int));
        table.Columns.Add("Ignored", typeof(string));
        table.Columns.Add("Text", typeof(string));
        table.Rows.Add(1, "secret", "skip");
        table.Rows.Add(2, "secret", "a\tb,\"c\nnext");
        table.Rows.Add(3, "secret", "=1+1");
        table.Rows.Add(4, "secret", DBNull.Value);
        table.Rows.Add(5, "secret", "outside");
        return table;
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ReorderedColumnsUseSequentialReadsAndStopExactlyAtLimit(bool tsv)
    {
        using var table = Table();
        using var reader = new TrackingReader(table.CreateDataReader());
        using var output = new NonSeekableOutput();
        var options = new DataExportOptions
        {
            Columns = new[] { "Text", "Id" }, CsvDelimiter = ';', RowOffset = 1, MaxRows = 3
        };
        if (tsv)
            await DataExporter.WriteTsvAsync(reader, output, options);
        else
            await DataExporter.WriteCsvAsync(reader, output, options);

        var separator = tsv ? "\t" : ";";
        var nl = Environment.NewLine;
        Assert.Equal("Text" + separator + "Id" + nl +
            "\"a\tb,\"\"c\nnext\"" + separator + "2" + nl +
            "'=1+1" + separator + "3" + nl +
            separator + "4" + nl, Encoding.UTF8.GetString(output.ToArray()));
        Assert.Equal(4, reader.ReadCalls);
        Assert.Equal(new[] { (2, 0), (2, 2), (3, 0), (3, 2), (4, 0), (4, 2) }, reader.NullChecks);
        Assert.Equal(new[] { (2, 0), (2, 2), (3, 0), (3, 2), (4, 0) }, reader.ValueReads);
        Assert.True(output.CanWrite);
        Assert.False(reader.IsClosed);
        Assert.Equal(';', options.CsvDelimiter);
        Assert.True(await reader.ReadAsync());
        Assert.Equal(5, reader.GetValue(0));
    }

    [Theory]
    [InlineData(0, 0, 0)]
    [InlineData(long.MaxValue, 1, 6)]
    public async Task EmptyWindowWritesSchemaAndAvoidsFieldReads(long offset, long limit, int expectedReads)
    {
        using var table = Table();
        using var reader = new TrackingReader(table.CreateDataReader());
        using var output = new MemoryStream();
        await DataExporter.WriteCsvAsync(reader, output, new DataExportOptions { RowOffset = offset, MaxRows = limit });
        Assert.Equal("Id,Ignored,Text" + Environment.NewLine, Encoding.UTF8.GetString(output.ToArray()));
        Assert.Equal(expectedReads, reader.ReadCalls);
        Assert.Empty(reader.NullChecks);
        Assert.Empty(reader.ValueReads);
    }

    [Fact]
    public async Task EmptyResultStillExportsHeaders()
    {
        using var table = Table();
        table.Clear();
        using var reader = new TrackingReader(table.CreateDataReader());
        using var output = new MemoryStream();
        await DataExporter.WriteTsvAsync(reader, output);
        Assert.Equal("Id\tIgnored\tText" + Environment.NewLine, Encoding.UTF8.GetString(output.ToArray()));
        Assert.Equal(1, reader.ReadCalls);
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("duplicate")]
    [InlineData("blank")]
    [InlineData("duplicate-schema")]
    [InlineData("blank-schema")]
    [InlineData("negative-offset")]
    [InlineData("negative-limit")]
    public async Task InvalidOptionsAndSchemaFailBeforeOutputOrReads(string kind)
    {
        using var table = Table();
        using var reader = new TrackingReader(table.CreateDataReader());
        using var output = new MemoryStream();
        if (kind == "duplicate-schema") reader.NameOverride = _ => "Same";
        if (kind == "blank-schema") reader.NameOverride = _ => "";
        var options = new DataExportOptions
        {
            Columns = kind switch
            {
                "missing" => new[] { "Missing" },
                "duplicate" => new[] { "Id", "Id" },
                "blank" => new[] { "" },
                _ => null
            },
            RowOffset = kind == "negative-offset" ? -1 : 0,
            MaxRows = kind == "negative-limit" ? -1 : null
        };
        await Assert.ThrowsAnyAsync<ArgumentException>(() => DataExporter.WriteCsvAsync(reader, output, options));
        Assert.Equal(0, reader.ReadCalls);
        Assert.Equal(0, output.Length);
        Assert.False(reader.IsClosed);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CancellationDuringReadStopsEvenWhenProviderIgnoresToken(bool duringNullCheck)
    {
        using var table = Table();
        using var reader = new TrackingReader(table.CreateDataReader());
        using var output = new MemoryStream();
        using var cancel = new CancellationTokenSource();
        if (duringNullCheck)
            reader.AfterNullCheck = () => cancel.Cancel();
        else
            reader.AfterRead = () => { if (reader.ReadCalls == 2) cancel.Cancel(); };
        var task = DataExporter.WriteCsvAsync(reader, output, cancellationToken: cancel.Token);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => task);
        Assert.True(task.IsCanceled);
        Assert.Equal(duringNullCheck ? 1 : 2, reader.ReadCalls);
        if (duringNullCheck) Assert.Empty(reader.ValueReads);
        else Assert.All(reader.ValueReads, x => Assert.Equal(1, x.Row));
        Assert.Equal(cancel.Token, reader.LastReadToken);
        Assert.False(reader.IsClosed);
        Assert.True(output.CanWrite);
    }

    [Fact]
    public async Task PreCancelledAndClosedReaderDoNotWrite()
    {
        using var table = Table();
        using var reader = new TrackingReader(table.CreateDataReader());
        using var output = new MemoryStream();
        using var cancel = new CancellationTokenSource();
        cancel.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => DataExporter.WriteTsvAsync(reader, output, cancellationToken: cancel.Token));
        Assert.False(reader.IsClosed);
        reader.Close();
        await Assert.ThrowsAsync<InvalidOperationException>(() => DataExporter.WriteCsvAsync(reader, output));
        Assert.Equal(0, reader.ReadCalls);
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task EncodingAndFormulaOptOutArePreserved()
    {
        using var table = Table();
        using var reader = new TrackingReader(table.CreateDataReader());
        using var output = new MemoryStream();
        await DataExporter.WriteTsvAsync(reader, output, new DataExportOptions
        {
            Columns = new[] { "Id", "Text" }, RowOffset = 2, MaxRows = 1,
            CsvIncludeHeaders = false, CsvProtectFormulas = false, CsvEncoding = Encoding.Unicode
        });
        Assert.Equal("3\t=1+1" + Environment.NewLine, Encoding.Unicode.GetString(output.ToArray()).TrimStart('\uFEFF'));
    }

    [Fact]
    public async Task XlsxSkipsUnselectedFieldsAndStopsAtLimit()
    {
        using var table = Table();
        using var reader = new TrackingReader(table.CreateDataReader());
        using var output = new MemoryStream();
        await DataExporter.WriteExcelStreamingAsync(reader, output, new StreamingExcelOptions
        {
            Columns = new[] { "Text", "Id" }, RowOffset = 1, MaxRows = 3
        });
        Assert.Equal(4, reader.ReadCalls);
        Assert.Equal(new[] { (2, 0), (2, 2), (3, 0), (3, 2), (4, 0), (4, 2) }, reader.NullChecks);
        Assert.Equal(new[] { (2, 0), (2, 2), (3, 0), (3, 2), (4, 0) }, reader.ValueReads);
        Assert.False(reader.IsClosed);
        output.Position = 0;
        using var workbook = new ClosedXML.Excel.XLWorkbook(output);
        var sheet = workbook.Worksheet(1);
        Assert.Equal("Text", sheet.Cell(1, 1).GetString());
        Assert.Equal(2, sheet.Cell(2, 2).GetValue<int>());
        Assert.Equal("=1+1", sheet.Cell(3, 1).GetString());
        Assert.False(sheet.Cell(3, 1).HasFormula);
    }

    [Fact]
    public async Task XlsxCancellationDuringNullCheckStopsBeforeValueRead()
    {
        using var table = Table();
        using var reader = new TrackingReader(table.CreateDataReader());
        using var output = new MemoryStream();
        using var cancel = new CancellationTokenSource();
        reader.AfterNullCheck = () => cancel.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            DataExporter.WriteExcelStreamingAsync(reader, output, cancellationToken: cancel.Token));
        Assert.Equal(1, reader.ReadCalls);
        Assert.Single(reader.NullChecks);
        Assert.Empty(reader.ValueReads);
        Assert.True(output.CanWrite);
        Assert.False(reader.IsClosed);
    }

    private sealed class NonSeekableOutput : MemoryStream
    {
        public override bool CanSeek => false;
        public override long Seek(long offset, SeekOrigin loc) => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
    }

    // Deliberately strict ordinal access; async methods ignore cancellation to
    // verify the exporter itself checks the token after provider calls.
    private sealed class TrackingReader(DbDataReader inner) : DbDataReader
    {
        private int lastOrdinal = -1;
        public int ReadCalls { get; private set; }
        public CancellationToken LastReadToken { get; private set; }
        public List<(int Row, int Ordinal)> NullChecks { get; } = new();
        public List<(int Row, int Ordinal)> ValueReads { get; } = new();
        public Action? AfterRead { get; set; }
        public Action? AfterNullCheck { get; set; }
        public Func<int, string>? NameOverride { get; set; }

        public override bool Read()
        {
            ReadCalls++;
            lastOrdinal = -1;
            return inner.Read();
        }
        public override Task<bool> ReadAsync(CancellationToken cancellationToken)
        {
            LastReadToken = cancellationToken;
            var result = Read();
            AfterRead?.Invoke();
            return Task.FromResult(result);
        }
        private void CheckOrdinal(int ordinal)
        {
            if (ordinal < lastOrdinal) throw new InvalidOperationException("Out-of-order field access");
            lastOrdinal = ordinal;
        }
        public override bool IsDBNull(int ordinal)
        {
            CheckOrdinal(ordinal);
            NullChecks.Add((ReadCalls, ordinal));
            return inner.IsDBNull(ordinal);
        }
        public override Task<bool> IsDBNullAsync(int ordinal, CancellationToken cancellationToken)
        {
            var result = IsDBNull(ordinal);
            AfterNullCheck?.Invoke();
            return Task.FromResult(result);
        }
        public override object GetValue(int ordinal)
        {
            CheckOrdinal(ordinal);
            ValueReads.Add((ReadCalls, ordinal));
            return inner.GetValue(ordinal);
        }

        public override int FieldCount => inner.FieldCount;
        public override bool HasRows => inner.HasRows;
        public override bool IsClosed => inner.IsClosed;
        public override int RecordsAffected => inner.RecordsAffected;
        public override int Depth => inner.Depth;
        public override object this[int ordinal] => GetValue(ordinal);
        public override object this[string name] => GetValue(GetOrdinal(name));
        public override string GetName(int ordinal) => NameOverride?.Invoke(ordinal) ?? inner.GetName(ordinal);
        public override int GetOrdinal(string name) => inner.GetOrdinal(name);
        public override string GetDataTypeName(int ordinal) => inner.GetDataTypeName(ordinal);
        public override Type GetFieldType(int ordinal) => inner.GetFieldType(ordinal);
        public override bool NextResult() => inner.NextResult();
        public override void Close() => inner.Close();
        protected override void Dispose(bool disposing)
        {
            if (disposing) inner.Dispose();
            base.Dispose(disposing);
        }
        public override IEnumerator GetEnumerator() => ((IEnumerable)inner).GetEnumerator();
        public override int GetValues(object[] values) => inner.GetValues(values);
        public override bool GetBoolean(int ordinal) => inner.GetBoolean(ordinal);
        public override byte GetByte(int ordinal) => inner.GetByte(ordinal);
        public override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length) => inner.GetBytes(ordinal, dataOffset, buffer, bufferOffset, length);
        public override char GetChar(int ordinal) => inner.GetChar(ordinal);
        public override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length) => inner.GetChars(ordinal, dataOffset, buffer, bufferOffset, length);
        public override Guid GetGuid(int ordinal) => inner.GetGuid(ordinal);
        public override short GetInt16(int ordinal) => inner.GetInt16(ordinal);
        public override int GetInt32(int ordinal) => inner.GetInt32(ordinal);
        public override long GetInt64(int ordinal) => inner.GetInt64(ordinal);
        public override float GetFloat(int ordinal) => inner.GetFloat(ordinal);
        public override double GetDouble(int ordinal) => inner.GetDouble(ordinal);
        public override string GetString(int ordinal) => inner.GetString(ordinal);
        public override decimal GetDecimal(int ordinal) => inner.GetDecimal(ordinal);
        public override DateTime GetDateTime(int ordinal) => inner.GetDateTime(ordinal);
    }
}
