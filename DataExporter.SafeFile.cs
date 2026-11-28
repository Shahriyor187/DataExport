namespace DataExportPackage;

public sealed partial class DataExporter
{
    /// <summary>
    /// Writes to a private staging file, validates the final byte length and publishes by rename.
    /// The destination remains unchanged if rendering, size validation or cancellation fails before publication.
    /// The optional positive size limit is a final-file check, not a temporary storage or memory quota.
    /// </summary>
    public Task<ExportFileResult> ExportFileAsync(string outputFile, ExportFormat format,
        DataExportOptions? options = null, long? maxOutputBytes = null,
        CancellationToken cancellationToken = default)
    {
        if (!Enum.IsDefined(format))
            throw new ArgumentOutOfRangeException(nameof(format));
        return SafeFileOutput.WriteAsync(outputFile, maxOutputBytes,
            (output, token) => ExportAsync(output, format, options, token), cancellationToken);
    }
}
