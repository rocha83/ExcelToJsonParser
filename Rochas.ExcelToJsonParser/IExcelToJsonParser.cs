using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Rochas.ExcelToJson
{
    public interface IExcelToJsonParser : IDisposable
    {
        string GetJsonStringFromTabular(string fileName, int skipRows = 0, string[] replaceFrom = null, string[] replaceTo = null, string[] headerColumns = null, bool onlySampleRow = false);
        string GetJsonStringFromTabular(Stream fileContent, int skipRows = 0, string[] replaceFrom = null, string[] replaceTo = null, string[] headerColumns = null, bool onlySampleRow = false);
        IEnumerable<object> GetJsonObjectFromTabular(string fileName, int skipRows = 0, string[] replaceFrom = null, string[] replaceTo = null, string[] headerColumns = null, bool onlySampleRow = false);
        string GetClassModelFromTabular(string fileName, int skipRows = 0, string[] replaceFrom = null, string[] replaceTo = null, string[] headerColumns = null);
        DataTable GetDataTable(string fileName, int skipRows = 0, bool useHeader = true);
        DataTable GetDataTable(Stream fileContent, int skipRows = 0, bool useHeader = true);
        IAsyncEnumerable<IDictionary<string, object>> StreamFromTabular(Stream fileContent, int skipRows = 0, string[] replaceFrom = null, string[] replaceTo = null, string[] headerColumns = null, CancellationToken cancellationToken = default);
        IDictionary<string, string> GetJsonStringsFromAllSheets(string fileName, int skipRows = 0, string[] replaceFrom = null, string[] replaceTo = null, string[] headerColumns = null);
        DataTable GetDataTable(string fileName, string sheetName);
        ValidationResult ValidateTabular(Stream fileContent, ValidationRule[] rules, int skipRows = 0);
        string ApplyTransforms(string value, TransformConfig[] transforms);
        byte[] TabularToExcel(IEnumerable<IDictionary<string, object>> data, string sheetName = "Sheet1", bool hasHeaders = true);
        byte[] CsvToExcel(Stream csvContent, string delimiter = ",", string sheetName = "Sheet1");
        byte[] CsvToExcel(string csvContent, string delimiter = ",", string sheetName = "Sheet1");
        byte[] JsonToExcel(string jsonContent, string sheetName = "Sheet1", bool hasHeaders = true);
        byte[] JsonToExcel(Stream jsonContent, string sheetName = "Sheet1", bool hasHeaders = true);
        byte[] JsonToCsv(string jsonContent, string delimiter = ",");
        Stream JsonToCsvStream(string jsonContent, string delimiter = ",");
        byte[] XmlToExcel(string xmlContent, string sheetName = "Sheet1", bool hasHeaders = true);
        byte[] XmlToExcel(Stream xmlContent, string sheetName = "Sheet1", bool hasHeaders = true);
        byte[] XmlToCsv(string xmlContent, string delimiter = ",");
        string GetJsonStringFromForm(string fileName, string sheetName, string[] replaceFrom = null, string[] replaceTo = null, string[] fieldNames = null);
        string GetJsonStringFromForm(Stream fileContent, string sheetName, string[] replaceFrom = null, string[] replaceTo = null, string[] fieldNames = null);
        object GetJsonObjectFromForm(string fileName, string sheetName, string[] replaceFrom = null, string[] replaceTo = null, string[] fieldNames = null);
        object GetJsonObjectFromForm(Stream fileContent, string sheetName, string[] replaceFrom = null, string[] replaceTo = null, string[] fieldNames = null);
        string GetClassModelFromForm(string fileName, string sheetName, string[] replaceFrom = null, string[] replaceTo = null, string[] fieldNames = null);
        IDictionary<string, object> GetDictionary(string fileName, string sheetName, string[] replaceFrom = null, string[] replaceTo = null, string[] fieldNames = null);
        IDictionary<string, object> GetDictionary(Stream fileContent, string sheetName, string[] replaceFrom = null, string[] replaceTo = null, string[] fieldNames = null);
    }
}
