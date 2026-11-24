using DataExportPackage;

using System.Text;
using Xunit;

namespace DataExport.Tests;

public sealed class FormulaWhitespaceTests
{
    [Theory]
    [InlineData("  =1+1", true)]
    [InlineData("\t+1", true)]
    [InlineData("\u00a0-1", true)]
    [InlineData(" \t@SUM(A1)", true)]
    [InlineData("  ordinary", false)]
    [InlineData("   ", false)]
    [InlineData("text=1+1", false)]
    public async Task DelimitedProtectionPreservesOriginalWhitespace(string value, bool formulaLike)
    {
        var exporter = DataExporter.FromObjects(new[] { new { Id = 1, Text = value } });
        foreach (var protect in new[] { true, false })
        foreach (var tsv in new[] { false, true })
        {
            using var output = new MemoryStream();
            var options = new DataExportOptions { CsvProtectFormulas = protect, Columns = new[] { "Id", "Text" } };
            if (tsv) await exporter.ToTsvAsync(output, options);
            else await exporter.ToCsvAsync(output, options);
            var delimiter = tsv ? '\t' : ',';
            var field = protect && formulaLike ? "'" + value : value;
            if (field.Contains(delimiter))
                field = "\"" + field + "\"";
            var expected = $"Id{delimiter}Text{Environment.NewLine}1{delimiter}{field}{Environment.NewLine}";
            Assert.Equal(expected, Encoding.UTF8.GetString(output.ToArray()));
        }
    }
}
