using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.IO;
using System.Text;
using Newtonsoft.Json;
using ExcelDataReader;
using ClosedXML.Excel;
using NJsonSchema.CodeGeneration.CSharp;

using Rochas.ExcelToJson.Helpers;

namespace Rochas.ExcelToJson
{
    public class ExcelToJsonParser : IDisposable
    {
        #region Tabular Sheet Parser Public Methods

        public string GetJsonStringFromTabular(string fileName, int skipRows = 0, string[] replaceFrom = null, string[] replaceTo = null, string[] headerColumns = null, bool onlySampleRow = false)
        {
            using (var fileContent = TabularSheetReader.GetFileStream(fileName))
            {
                return GetJsonStringFromTabular(fileContent, skipRows, replaceFrom, replaceTo, headerColumns, onlySampleRow);
            }
        }

        public string GetJsonStringFromTabular(Stream fileContent, int skipRows = 0, string[] replaceFrom = null, string[] replaceTo = null, string[] headerColumns = null, bool onlySampleRow = false)
        {
            var counter = 0;
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            using (var result = new StringWriter())
            {
                var readerConfig = new ExcelReaderConfiguration()
                {
                    FallbackEncoding = Encoding.GetEncoding(1252)
                };

                using (var reader = ExcelReaderFactory.CreateReader(fileContent, readerConfig))
                {
                    using (var writer = new JsonTextWriter(result))
                    {
                        writer.Formatting = Formatting.Indented;
                        writer.WriteStartArray();

                        while (skipRows > 0)
                        {
                            reader.Read();
                            skipRows--;
                        }

                        reader.Read();

                        if (headerColumns == null)
                            headerColumns = TabularSheetReader.GetHeaderColumns(reader);
                        else
                        {
                            if (headerColumns.Length < reader.FieldCount)
                                throw new Exception("Invalid column amount");
                        }

                        TabularSheetConversor.ApplyColumnNamesReplace(headerColumns, replaceFrom, replaceTo);

                        do
                        {
                            while (reader.Read() && (!onlySampleRow || (onlySampleRow && counter < 1)))
                            {
                                JsonWriterHelper.WriteItemJsonBodyFromReader(reader, writer, headerColumns);
                                counter += 1;
                            }

                        } while (reader.NextResult());

                        writer.WriteEndArray();
                    }
                }

                return result.ToString();
            }
        }

        public IEnumerable<object> GetJsonObjectFromTabular(string fileName, int skipRows = 0, string[] replaceFrom = null, string[] replaceTo = null, string[] headerColumns = null, bool onlySampleRow = false)
        {
            var strJson = GetJsonStringFromTabular(fileName, skipRows, replaceFrom, replaceTo, headerColumns, onlySampleRow);
            return JsonConvert.DeserializeObject<IEnumerable<object>>(strJson);
        }

        public IEnumerable<object> GetJsonObjectFromTabular(Stream fileContent, int skipRows = 0, string[] replaceFrom = null, string[] replaceTo = null, string[] headerColumns = null, bool onlySampleRow = false)
        {
            var strJson = GetJsonStringFromTabular(fileContent, skipRows, replaceFrom, replaceTo, headerColumns, onlySampleRow);
            return JsonConvert.DeserializeObject<IEnumerable<object>>(strJson);
        }

        public string GetClassModelFromTabular(string fileName, int skipRows = 0, string[] replaceFrom = null, string[] replaceTo = null, string[] headerColumns = null)
        {
            string result = null;
            var jsonContent = GetJsonStringFromTabular(fileName, skipRows, replaceFrom, replaceTo, headerColumns, true);

            if (!string.IsNullOrWhiteSpace(jsonContent))
            {
                var schema = NJsonSchema.JsonSchema.FromSampleJson(jsonContent);

                var genOptions = new CSharpGeneratorSettings()
                {
                    GenerateDataAnnotations = false,
                    GenerateDefaultValues = false,
                    GenerateJsonMethods = true
                };

                var generator = new CSharpGenerator(schema, genOptions);
                var className = fileName.Replace(".xlsx", string.Empty).Replace(".xls", string.Empty);

                result = generator.GenerateFile(className);
            }

            return result;
        }

        #region DataTable results support

        public DataTable GetDataTable(string fileName, int skipRows = 0, bool useHeader = true)
        {
            DataTable result = null;

            if (!string.IsNullOrWhiteSpace(fileName))
            {
                using (var fileContent = TabularSheetReader.GetFileStream(fileName))
                {
                    result = TabularSheetReader.GetDataTable(fileContent, skipRows, useHeader);
                }
            }

            return result;
        }

        public static DataTable GetDataTable(Stream fileContent, int skipRows = 0, bool useHeader = true)
        {
            return TabularSheetReader.GetDataTable(fileContent, skipRows, useHeader);
        }

        #endregion

        #endregion

        #region Form Sheet Parser Public Methods

        public string GetJsonStringFromForm(string fileName, string sheetName, string[] replaceFrom = null, string[] replaceTo = null, string[] fieldNames = null)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                throw new Exception("File name not informed");

            using (var engine = new XLWorkbook(fileName))
            {
                var parsedData = FormSheetHelper.ParseFormSheet(engine, sheetName, fieldNames);
                return JsonWriterHelper.WriteJsonBodyFromNamedFields(parsedData);
            }
        }

        public static string GetJsonStringFromForm(Stream fileContent, string sheetName, string[] replaceFrom = null, string[] replaceTo = null, string[] fieldNames = null)
        {
            if (fileContent == null)
                throw new Exception("File content not informed");

            using (var engine = new XLWorkbook(fileContent))
            {
                var parsedData = FormSheetHelper.ParseFormSheet(engine, sheetName, fieldNames);
                return JsonWriterHelper.WriteJsonBodyFromNamedFields(parsedData);
            }
        }

        public object GetJsonObjectFromForm(string fileName, string sheetName, string[] replaceFrom = null, string[] replaceTo = null, string[] fieldNames = null)
        {
            var strJson = GetJsonStringFromForm(fileName, sheetName, replaceTo, fieldNames);
            return JsonConvert.DeserializeObject(strJson);
        }

        public object GetJsonObjectFromForm(Stream fileContent, string sheetName, string[] replaceFrom = null, string[] replaceTo = null, string[] fieldNames = null)
        {
            var strJson = GetJsonStringFromForm(fileContent, sheetName, replaceTo, fieldNames);
            return JsonConvert.DeserializeObject(strJson);
        }

        public string GetClassModelFromForm(string fileName, string sheetName, string[] replaceFrom = null, string[] replaceTo = null, string[] fieldNames = null)
        {
            string result = null;
            var jsonContent = GetJsonStringFromForm(fileName, sheetName, replaceFrom, replaceTo, fieldNames);

            if (!string.IsNullOrWhiteSpace(jsonContent))
            {
                var schema = NJsonSchema.JsonSchema.FromSampleJson(jsonContent);

                var genOptions = new CSharpGeneratorSettings()
                {
                    GenerateDataAnnotations = false,
                    GenerateDefaultValues = false,
                    GenerateJsonMethods = true
                };
                var generator = new CSharpGenerator(schema, genOptions);

                var className = fileName.Replace(".xlsx", string.Empty).Replace(".xls", string.Empty);

                result = generator.GenerateFile(className);
            }

            return result;
        }

        #region Dictionary results support

        public IDictionary<string, object> GetDictionary(string fileName, string sheetName, string[] replaceFrom = null, string[] replaceTo = null, string[] fieldNames = null)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                throw new Exception("File name not informed");

            using (var engine = new XLWorkbook(fileName))
            {
                return FormSheetHelper.ParseFormSheet(engine, sheetName, fieldNames);
            }
        }

        public IDictionary<string, object> GetDictionary(Stream fileContent, string sheetName, string[] replaceFrom = null, string[] replaceTo = null, string[] fieldNames = null)
        {
            if (fileContent == null)
                throw new Exception("File content not informed");

            using (var engine = new XLWorkbook(fileContent))
            {
                return FormSheetHelper.ParseFormSheet(engine, sheetName, fieldNames);
            }
        }

        #endregion

        #endregion

        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }
    }
}
