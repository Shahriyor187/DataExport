using ClosedXML.Excel;
using DataExportPackage;
using Xunit;

namespace DataExport.Tests;

public sealed class ExcelReportBuilderTests
{
    private sealed record Address(string City);
    private sealed record Student(string Name, int Score, Address Address);

    [Fact]
    public void BuildsMultipleSheetsWithNestedColumnsAndTotals()
    {
        using var output = new MemoryStream();
        using (var builder = new ExcelReportBuilder())
        {
            builder.AddSheet("Students",
                [new Student("Amina", 1500, new Address("Tashkent")),
                 new Student("Ali", 1400, new Address("Samarkand"))],
                new ExcelReportSheetOptions
                {
                    Columns =
                    [
                        new ExcelReportColumn("Name", "Student", Width: 25),
                        new ExcelReportColumn("Score", "SAT Score", "#,##0"),
                        new ExcelReportColumn("Address.City", "City")
                    ]
                });
            builder.AddTotals("Students", 2);
            builder.AddSheet("Summary",
                new[] { new { Metric = "Students", Value = 2 } });
            builder.SaveAs(output);
        }

        output.Position = 0;
        using var book = new XLWorkbook(output);
        Assert.Equal(2, book.Worksheets.Count);
        var students = book.Worksheet("Students");
        Assert.Equal("SAT Score", students.Cell("B1").GetString());
        Assert.Equal("Tashkent", students.Cell("C2").GetString());
        Assert.Equal(1500, students.Cell("B2").GetDouble());
        Assert.Equal("SUM(B2:B3)", students.Cell("B4").FormulaA1);
        Assert.Equal("Students", book.Worksheet("Summary").Cell("A1").GetString());
    }

    [Fact]
    public void RejectsInvalidTotalColumns()
    {
        using var builder = new ExcelReportBuilder();
        builder.AddSheet("Data", new[] { new { Name = "Amina", Score = 1450 } });
        Assert.Throws<ArgumentOutOfRangeException>(() => builder.AddTotals("Data", 9));
    }
}
