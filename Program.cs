using DataExportPackage;

class Program
{
    static void Main(string[] args)
    {
        // Sample data
        var data = new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    {"Name", "John Doe"},
                    {"Age", 30},
                    {"Country", "USA"}
                },
                new Dictionary<string, object>
                {
                    {"Name", "Jane Smith"},
                    {"Age", 25},
                    {"Country", "Canada"}
                }
            };

        // Create an instance of DataExporter
        DataExporter dataExporter = new DataExporter(data);

        string jsonFilePath = @"C:\YourFolder\data.json";
        string xmlFilePath = @"C:\YourFolder\data.xml";
        string excelFilePath = @"C:\YourFolder\data.xlsx";

        // Export data to JSON
        dataExporter.ToJson(jsonFilePath);

        // Export data to XML
        dataExporter.ToXml(xmlFilePath);

        // Export data to Excel
        dataExporter.ToExcel(excelFilePath);

        Console.WriteLine("Data exported successfully.");
    }
}