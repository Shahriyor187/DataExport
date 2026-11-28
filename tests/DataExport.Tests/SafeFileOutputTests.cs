using DataExportPackage;
using Xunit;

namespace DataExport.Tests;

public sealed class SafeFileOutputTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
    private readonly string destination;
    private readonly DataExporter exporter = DataExporter.FromObjects(new[] { new { Name = "Amina", Score = 1450 } });

    public SafeFileOutputTests()
    {
        Directory.CreateDirectory(directory);
        destination = Path.Combine(directory, "output.csv");
        File.WriteAllText(destination, "original");
        
    }

    private Task<ExportFileResult> Write(long? limit = null, CancellationToken token = default) =>
        exporter.ExportFileAsync(destination, ExportFormat.Csv, maxOutputBytes: limit, cancellationToken: token);

    [Fact]
    public async Task SuccessfulPublicationReportsActualFileSize()
    {
        var result = await Write();
        Assert.Equal(destination, result.OutputPath);
        Assert.Equal(new FileInfo(destination).Length, result.OutputBytes);
        Assert.Contains("Amina", File.ReadAllText(destination));
        Assert.True(result.Elapsed >= TimeSpan.Zero);
        Assert.Empty(Directory.GetDirectories(directory));
    }

    [Fact]
    public async Task OutputLimitPreservesExistingFile()
    {
        await Assert.ThrowsAsync<IOException>(() => Write(1));
        Assert.Equal("original", File.ReadAllText(destination));
        Assert.Empty(Directory.GetDirectories(directory));
    }

    [Fact]
    public async Task ExactOutputLimitIsAccepted()
    {
        var first = await Write();
        var second = await Write(first.OutputBytes);
        Assert.Equal(first.OutputBytes, second.OutputBytes);
    }

    [Fact]
    public async Task InvalidLimitPreservesExistingFile()
    {
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => Write(0));
        Assert.Equal("original", File.ReadAllText(destination));
        Assert.Empty(Directory.GetDirectories(directory));
    }

    [Fact]
    public async Task CancellationPreservesExistingFile()
    {
        using var source = new CancellationTokenSource();
        source.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => Write(token: source.Token));
        Assert.Equal("original", File.ReadAllText(destination));
        Assert.Empty(Directory.GetDirectories(directory));
    }

    [Fact]
    public async Task RenderFailurePreservesExistingFile()
    {
        await Assert.ThrowsAnyAsync<ArgumentException>(() => exporter.ExportFileAsync(
            destination, ExportFormat.Csv, new DataExportOptions { RowOffset = -1 }));
        Assert.Equal("original", File.ReadAllText(destination));
        Assert.Empty(Directory.GetDirectories(directory));
    }
    

    public void Dispose() => Directory.Delete(directory, recursive: true);
}
