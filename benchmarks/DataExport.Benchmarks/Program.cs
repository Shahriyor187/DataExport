using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Running;
using DataExportPackage;
using System.Diagnostics;
using System.IO.Compression;
using System.Runtime.CompilerServices;
using System.Text.Json;
using System.Xml;

if (args.Contains("--stress", StringComparer.OrdinalIgnoreCase))
{
    await StressRunner.RunAsync(args);
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(ExcelExportBenchmarks).Assembly).Run(args);

[MemoryDiagnoser]
public class ExcelExportBenchmarks
{
    [Params(1_000, 10_000)]
    public int RowCount { get; set; }

    [Benchmark(Baseline = true, Description = "ClosedXML / buffered")]
    public void BufferedExcel()
    {
        var data = Enumerable.Range(0, RowCount).Select(i =>
            new ReportRow(i, "Student " + i, 900 + i % 700, i % 2 == 0, new DateTime(2026, 1, 1).AddDays(i % 365))).ToArray();
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            DataExporter.FromObjects(data).ToExcel(path, new DataExportOptions
            {
                ExcelIncludeHeaders = true,
                ExcelAutoFitColumns = false
            });
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Benchmark(Description = "Open XML / forward-only")]
    public async Task StreamingExcel()
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            await using var stream = File.Create(path);
            await DataExporter.WriteExcelStreamingAsync(
                ReportRows.Generate(RowCount), stream,
                new StreamingExcelOptions { Columns = ReportRows.Columns });
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}

public sealed record ReportRow(int Id, string Name, int Score, bool Passed, DateTime TakenOn);

public static class ReportRows
{
    public static readonly string[] Columns = ["Id", "Name", "Score", "Passed", "TakenOn"];

    public static async IAsyncEnumerable<ReportRow> Generate(
        int count, [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        for (var i = 0; i < count; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            yield return new ReportRow(i, "Student " + i, 900 + i % 700,
                i % 2 == 0, new DateTime(2026, 1, 1).AddDays(i % 365));
        }
        await Task.CompletedTask;
    }
}

public static class StressRunner
{
    public static async Task RunAsync(string[] args)
    {
        var rowCount = GetInt(args, "--rows", 100_000);
        var maxRowsPerWorksheet = GetInt(args, "--max-rows-per-sheet", 1_048_576);
        var maxPeakMb = GetInt(args, "--max-peak-mb", 0);
        var resultFile = GetOption(args, "--results");
        if (rowCount < 0 || rowCount > 5_000_000)
            throw new ArgumentOutOfRangeException(nameof(rowCount), "Allow 0..5,000,000 rows.");
        if (maxPeakMb < 0)
            throw new ArgumentOutOfRangeException(nameof(maxPeakMb));

        var path = Path.Combine(Path.GetTempPath(), "dataexport-stress-" + Guid.NewGuid().ToString("N") + ".xlsx");
        var clock = Stopwatch.StartNew();
        long peakWorkingSet = Environment.WorkingSet;
        var startAlloc = GC.GetTotalAllocatedBytes(precise: true);
        using var samplerStop = new CancellationTokenSource();
        var sampler = Task.Run(async () =>
        {
            while (!samplerStop.IsCancellationRequested)
            {
                var value = Process.GetCurrentProcess().WorkingSet64;
                if (value > Interlocked.Read(ref peakWorkingSet))
                    Interlocked.Exchange(ref peakWorkingSet, value);
                try { await Task.Delay(50, samplerStop.Token); }
                catch (OperationCanceledException) { break; }
            }
        });
        try
        {
            await using (var output = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite,
                FileShare.None, 64 * 1024, FileOptions.SequentialScan))
            {
                await DataExporter.WriteExcelStreamingAsync(ReportRows.Generate(rowCount),
                    output, new StreamingExcelOptions
                    {
                        Columns = ReportRows.Columns,
                        MaxRowsPerWorksheet = maxRowsPerWorksheet,
                        SplitLargeWorksheets = true
                    });
            }

            clock.Stop();
            samplerStop.Cancel();
            await sampler;
            peakWorkingSet = Math.Max(peakWorkingSet, Process.GetCurrentProcess().PeakWorkingSet64);
            var allocatedBytes = GC.GetTotalAllocatedBytes(precise: true) - startAlloc;
            var bytesOnDisk = new FileInfo(path).Length;

            // Read every row using XmlReader to verify without loading the workbook.
            var sheetRowCounts = VerifyWorksheets(path, rowCount, maxRowsPerWorksheet);
            var summary = new
            {
                rows = rowCount,
                dataColumns = ReportRows.Columns.Length,
                sheetRowCounts,
                elapsedSeconds = Math.Round(clock.Elapsed.TotalSeconds, 3),
                rowsPerSecond = clock.Elapsed.TotalSeconds > 0
                    ? Math.Round(rowCount / clock.Elapsed.TotalSeconds, 0) : 0,
                allocatedMb = Math.Round(allocatedBytes / 1024d / 1024d, 2),
                processPeakWorkingSetMb = Math.Round(peakWorkingSet / 1024d / 1024d, 2),
                outputMb = Math.Round(bytesOnDisk / 1024d / 1024d, 2),
                dotnet = Environment.Version.ToString(),
                processorCount = Environment.ProcessorCount
            };
            var json = JsonSerializer.Serialize(summary, new JsonSerializerOptions { WriteIndented = true });
            Console.WriteLine(json);
            if (!string.IsNullOrWhiteSpace(resultFile))
            {
                var directory = Path.GetDirectoryName(Path.GetFullPath(resultFile));
                if (!string.IsNullOrEmpty(directory))
                    Directory.CreateDirectory(directory);
                await File.WriteAllTextAsync(resultFile, json);
            }
            if (maxPeakMb != 0 && peakWorkingSet > (long)maxPeakMb * 1024 * 1024)
                throw new InvalidOperationException(
                    $"Peak process working set {summary.processPeakWorkingSetMb} MB exceeds configured {maxPeakMb} MB.");
        }
        finally
        {
            samplerStop.Cancel();
            await sampler;
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static int[] VerifyWorksheets(string file, int recordCount, int maxRows)
    {
        var counts = new List<int>();
        using var zip = ZipFile.OpenRead(file);
        var entries = zip.Entries.Where(e => e.FullName.StartsWith("xl/worksheets/sheet", StringComparison.Ordinal)
            && e.FullName.EndsWith(".xml", StringComparison.Ordinal)).OrderBy(e => e.FullName, StringComparer.Ordinal);
        foreach (var sheet in entries)
        {
            int rows = 0;
            using var stream = sheet.Open();
            using var xml = XmlReader.Create(stream, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null
            });
            while (xml.Read())
                if (xml.NodeType == XmlNodeType.Element && xml.LocalName == "row"
                    && xml.NamespaceURI == "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
                    rows++;
            counts.Add(rows);
        }
        int expectedSheets = recordCount == 0 ? 1 :
            (int)Math.Ceiling((double)recordCount / (maxRows - 1));
        if (counts.Count != expectedSheets || counts.Sum() != recordCount + counts.Count ||
            counts.Any(count => count > maxRows))
            throw new InvalidDataException(
                $"XLSX verification failed. Expected {recordCount} data rows across {expectedSheets} sheets; got [{string.Join(", ", counts)}].");
        return counts.ToArray();
    }

    private static string? GetOption(string[] args, string name)
    {
        var index = Array.IndexOf(args, name);
        if (index < 0) return null;
        if (index + 1 >= args.Length) throw new ArgumentException("Missing argument for " + name);
        return args[index + 1];
    }

    private static int GetInt(string[] args, string name, int defaultValue)
    {
        var value = GetOption(args, name);
        return value is null ? defaultValue : int.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
    }
}
