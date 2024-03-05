using ClosedXML.Excel;
using Newtonsoft.Json;
using System.Xml;

namespace DataExportPackage;
public class DataExporter
{
    private List<Dictionary<string, object>> data;
    public DataExporter(List<Dictionary<string, object>> data)
    {
        this.data = data;
    }
    public void ToJson(string filePath)
    {
        string jsonData = JsonConvert.SerializeObject(data, Newtonsoft.Json.Formatting.Indented);
        File.WriteAllText(filePath, jsonData);
    }
    public void ToXml(string filePath)
    {
        XmlDocument xmlDoc = new XmlDocument();
        XmlDeclaration xmlDeclaration = xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", null);
        XmlElement root = xmlDoc.DocumentElement;
        xmlDoc.InsertBefore(xmlDeclaration, root);

        XmlElement dataElement = xmlDoc.CreateElement(string.Empty, "data", string.Empty);
        xmlDoc.AppendChild(dataElement);

        foreach (var item in data)
        {
            XmlElement entryElement = xmlDoc.CreateElement(string.Empty, "entry", string.Empty);
            dataElement.AppendChild(entryElement);

            foreach (var kvp in item)
            {
                XmlElement element = xmlDoc.CreateElement(string.Empty, kvp.Key, string.Empty);
                element.InnerText = kvp.Value.ToString();
                entryElement.AppendChild(element);
            }
        }

        xmlDoc.Save(filePath);
    }

    public void ToExcel(string filePath)
    {
        var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Data");

        int row = 1;
        foreach (var item in data)
        {
            int col = 1;
            foreach (var kvp in item)
            {
                ws.Cell(row, col).Value = kvp.Value.ToString();
                col++;
            }
            row++;
        }

        wb.SaveAs(filePath);
    }
}