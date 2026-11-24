using DataExportPackage;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Text.Json;

internal static class DelimitedStress
{
    private sealed record Row(int Id, string Text, string Formula);
    private static async IAsyncEnumerable<Row> Rows(int count, [EnumeratorCancellation] CancellationToken token = default)
    {
        for (int i = 0; i < count; i++)
        {
            token.ThrowIfCancellationRequested();
            yield return new Row(i, "Row " + i, "=1+1");
        }
        await Task.CompletedTask;
    }

    internal static async Task RunAsync(string[] args)
    {
        if (args.Length != 3 || args[1] is not ("csv" or "tsv") ||
            !int.TryParse(args[2], out var rows) || rows < 1 || rows > 5_000_000)
            throw new ArgumentException("Usage: --delimited-stress <csv|tsv> <rows:1..5000000>");
        var format = args[1];
        var separator = format == "csv" ? "," : "\t";
        var directory = Directory.CreateTempSubdirectory("dataexport-bench-");
        try
        {
            var path = Path.Combine(directory.FullName, "output." + format);
            var allocatedBefore = GC.GetTotalAllocatedBytes(true);
            var timer = Stopwatch.StartNew();
            await using (var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write,
                FileShare.None, 65536, useAsync: true))
            {
                var options = new DataExportOptions
                {
                    Columns = new[] { "Id", "Text", "Formula" },
                    CsvProtectFormulas = true
                };
                if (format == "csv") await DataExporter.WriteCsvAsync(Rows(rows), stream, options);
                else await DataExporter.WriteTsvAsync(Rows(rows), stream, options);
            }
            timer.Stop();
            var allocated = GC.GetTotalAllocatedBytes(true) - allocatedBefore;
            using var process = Process.GetCurrentProcess();
            var peak = process.PeakWorkingSet64;
            var size = new FileInfo(path).Length;

            // Fixture contains no newlines or delimiters in individual values.
            // Verify every record in order without materializing the output.
            using var reader = File.OpenText(path);
            if (await reader.ReadLineAsync() != string.Join(separator, "Id", "Text", "Formula"))
                throw new InvalidDataException("Invalid header.");
            for (int i = 0; i < rows; i++)
                if (await reader.ReadLineAsync() != $"{i}{separator}Row {i}{separator}'=1+1")
                    throw new InvalidDataException($"Incorrect record at index {i}.");
            if (await reader.ReadLineAsync() is not null)
                throw new InvalidDataException("Unexpected extra output.");
            Console.WriteLine("DELIMITED_STRESS " + JsonSerializer.Serialize(new
            {
                format, rows, elapsedSeconds = timer.Elapsed.TotalSeconds,
                rowsPerSecond = rows / timer.Elapsed.TotalSeconds,
                allocatedBytes = allocated, peakWorkingSetBytes = peak,
                outputBytes = size, verified = true,
                runtime = Environment.Version.ToString()
            }));
        }
        finally { directory.Delete(true); }
    }
}
