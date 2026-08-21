using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using ExcelDataReader;
using ClosedXML.Excel;
using NJsonSchema.CodeGeneration.CSharp;

namespace Rochas.ExcelToJson
{
    public class ExcelToJsonParser : IExcelToJsonParser, IDisposable
    {
        #region Tabular Sheet Parser

        public string GetJsonStringFromTabular(string fileName, int skipRows = 0, string[] replaceFrom = null, string[] replaceTo = null, string[] headerColumns = null, bool onlySampleRow = false)
        {
            using (var fileContent = GetFileStream(fileName))
            {
                return GetJsonStringFromTabular(fileContent, skipRows, replaceFrom, replaceTo, headerColumns, onlySampleRow);
            }
        }

        public string GetJsonStringFromTabular(Stream fileContent, int skipRows = 0, string[] replaceFrom = null, string[] replaceTo = null, string[] headerColumns = null, bool onlySampleRow = false)
        {
            var counter = 0;
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            using (var ms = new MemoryStream())
            {
                var readerConfig = new ExcelReaderConfiguration()
                {
                    FallbackEncoding = Encoding.GetEncoding(1252)
                };
                using (var reader = ExcelReaderFactory.CreateReader(fileContent, readerConfig))
                {
                    using (var writer = new Utf8JsonWriter(ms, new JsonWriterOptions { Indented = true }))
                    {
                        writer.WriteStartArray();

                        while (skipRows > 0)
                        {
                            reader.Read();
                            skipRows--;
                        }

                        reader.Read();

                        if (headerColumns == null)
                            headerColumns = GetHeaderColumns(reader);
                        else
                        {
                            if (headerColumns.Length < reader.FieldCount)
                                throw new Exception("Invalid column amount");
                        }

                        ApplyColumnNamesReplace(headerColumns, replaceFrom, replaceTo);

                        do
                        {
                            while (reader.Read() && (!onlySampleRow || (onlySampleRow && counter < 1)))
                            {
                                WriteItemJsonBodyFromReader(reader, writer, headerColumns);
                                counter += 1;
                            }

                        } while (reader.NextResult());

                        writer.WriteEndArray();
                    }
                }

                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }

        public IEnumerable<object> GetJsonObjectFromTabular(string fileName, int skipRows = 0, string[] replaceFrom = null, string[] replaceTo = null, string[] headerColumns = null, bool onlySampleRow = false)
        {
            var strJson = GetJsonStringFromTabular(fileName, skipRows, replaceFrom, replaceTo, headerColumns, onlySampleRow);
            return JsonSerializer.Deserialize<IEnumerable<object>>(strJson);
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
                var className = Path.GetFileNameWithoutExtension(fileName);
                result = generator.GenerateFile(className);
            }
            return result;
        }

        public DataTable GetDataTable(string fileName, int skipRows = 0, bool useHeader = true)
        {
            if (string.IsNullOrWhiteSpace(fileName)) return null;
            using (var fileContent = GetFileStream(fileName))
            {
                return GetDataTable(fileContent, skipRows, useHeader);
            }
        }

        public DataTable GetDataTable(Stream fileContent, int skipRows = 0, bool useHeader = true)
        {
            if (fileContent == null) return null;

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var readerConfig = new ExcelReaderConfiguration()
            {
                FallbackEncoding = Encoding.GetEncoding(1252)
            };
            var reader = ExcelReaderFactory.CreateReader(fileContent, readerConfig);
            var config = new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = useHeader }
            };

            while (skipRows > 0) { reader.Read(); skipRows--; }

            return reader.AsDataSet(config).Tables[0];
        }

        #endregion

        #region Streaming

        public async IAsyncEnumerable<IDictionary<string, object>> StreamFromTabular(
            Stream fileContent,
            int skipRows = 0,
            string[] replaceFrom = null,
            string[] replaceTo = null,
            string[] headerColumns = null,
            [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken cancellationToken = default)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var readerConfig = new ExcelReaderConfiguration()
            {
                FallbackEncoding = Encoding.GetEncoding(1252)
            };

            using (var reader = ExcelReaderFactory.CreateReader(fileContent, readerConfig))
            {
                while (skipRows > 0) { reader.Read(); skipRows--; }
                reader.Read();

                if (headerColumns == null)
                    headerColumns = GetHeaderColumns(reader);
                else if (headerColumns.Length < reader.FieldCount)
                    throw new Exception("Invalid column amount");

                ApplyColumnNamesReplace(headerColumns, replaceFrom, replaceTo);

                do
                {
                    while (reader.Read())
                    {
                        if (cancellationToken.IsCancellationRequested) yield break;
                        var row = new Dictionary<string, object>();
                        for (var col = 0; col < headerColumns.Length; col++)
                            row[headerColumns[col]] = reader.GetValue(col);
                        yield return row;
                    }
                } while (reader.NextResult());
            }
        }

        #endregion

        #region Multi-sheet

        public IDictionary<string, string> GetJsonStringsFromAllSheets(string fileName, int skipRows = 0, string[] replaceFrom = null, string[] replaceTo = null, string[] headerColumns = null)
        {
            var result = new Dictionary<string, string>();
            using (var fileContent = GetFileStream(fileName))
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                var readerConfig = new ExcelReaderConfiguration()
                {
                    FallbackEncoding = Encoding.GetEncoding(1252)
                };

                using (var reader = ExcelReaderFactory.CreateReader(fileContent, readerConfig))
                {
                    var sheetIndex = 0;
                    do
                    {
                        var sheetName = reader.Name ?? $"Sheet{sheetIndex + 1}";
                        var sheetSkip = skipRows;

                        using (var sheetMs = new MemoryStream())
                        {
                            using (var writer = new Utf8JsonWriter(sheetMs, new JsonWriterOptions { Indented = true }))
                            {
                                writer.WriteStartArray();
                                while (sheetSkip > 0) { reader.Read(); sheetSkip--; }
                                reader.Read();

                                var headers = headerColumns ?? GetHeaderColumns(reader);
                                if (headerColumns == null)
                                    ApplyColumnNamesReplace(headers, replaceFrom, replaceTo);

                                do { while (reader.Read()) { WriteItemJsonBodyFromReader(reader, writer, headers); } } while (reader.NextResult());
                                writer.WriteEndArray();
                            }
                            result[sheetName] = Encoding.UTF8.GetString(sheetMs.ToArray());
                        }
                        sheetIndex++;
                    } while (reader.NextResult());
                }
            }
            return result;
        }

        public DataTable GetDataTable(string fileName, string sheetName)
        {
            using (var fileContent = GetFileStream(fileName))
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                var readerConfig = new ExcelReaderConfiguration()
                {
                    FallbackEncoding = Encoding.GetEncoding(1252)
                };
                var reader = ExcelReaderFactory.CreateReader(fileContent, readerConfig);
                var config = new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = true }
                };
                var dataSet = reader.AsDataSet(config);
                return dataSet.Tables.Contains(sheetName) ? dataSet.Tables[sheetName] : dataSet.Tables[0];
            }
        }

        #endregion

        #region Validation

        public ValidationResult ValidateTabular(Stream fileContent, ValidationRule[] rules, int skipRows = 0)
        {
            var result = new ValidationResult();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            var readerConfig = new ExcelReaderConfiguration()
            {
                FallbackEncoding = Encoding.GetEncoding(1252)
            };

            using (var reader = ExcelReaderFactory.CreateReader(fileContent, readerConfig))
            {
                while (skipRows > 0) { reader.Read(); skipRows--; }
                reader.Read();
                var headers = GetHeaderColumns(reader);

                var rowIndex = 0;
                do
                {
                    while (reader.Read())
                    {
                        rowIndex++;
                        var rowHasError = false;
                        for (var col = 0; col < headers.Length && col < reader.FieldCount; col++)
                        {
                            var colName = headers[col];
                            var rawValue = reader.GetValue(col)?.ToString() ?? string.Empty;
                            var applicableRules = rules.Where(r => string.Equals(r.ColumnName, colName, StringComparison.OrdinalIgnoreCase));
                            foreach (var rule in applicableRules)
                            {
                                var error = ValidateCell(rowIndex, colName, rawValue, rule);
                                if (error != null) { result.Errors.Add(error); rowHasError = true; }
                            }
                        }
                        result.TotalRows++;
                        if (rowHasError) result.ErrorRows++; else result.ValidRows++;
                    }
                } while (reader.NextResult());
            }

            result.IsValid = result.ErrorRows == 0;
            return result;
        }

        private static ValidationError ValidateCell(int rowIndex, string columnName, string rawValue, ValidationRule rule)
        {
            switch (rule.Type?.ToLower())
            {
                case "required":
                    if (string.IsNullOrWhiteSpace(rawValue))
                        return new ValidationError { RowNumber = rowIndex, ColumnName = columnName, ErrorType = "required", Message = rule.Message ?? $"Campo '{columnName}' é obrigatório", RawValue = rawValue };
                    break;
                case "max_length":
                    if (rule.MaxLength.HasValue && rawValue.Length > rule.MaxLength.Value)
                        return new ValidationError { RowNumber = rowIndex, ColumnName = columnName, ErrorType = "max_length", Message = rule.Message ?? $"Campo '{columnName}' excede {rule.MaxLength} caracteres", RawValue = rawValue };
                    break;
                case "min_length":
                    if (rule.MinLength.HasValue && rawValue.Length < rule.MinLength.Value)
                        return new ValidationError { RowNumber = rowIndex, ColumnName = columnName, ErrorType = "min_length", Message = rule.Message ?? $"Campo '{columnName}' requer no mínimo {rule.MinLength} caracteres", RawValue = rawValue };
                    break;
                case "numeric":
                    if (!string.IsNullOrWhiteSpace(rawValue) && !decimal.TryParse(rawValue, out _))
                        return new ValidationError { RowNumber = rowIndex, ColumnName = columnName, ErrorType = "numeric", Message = rule.Message ?? $"Campo '{columnName}' deve ser numérico", RawValue = rawValue };
                    break;
                case "date":
                    if (!string.IsNullOrWhiteSpace(rawValue) && !DateTime.TryParse(rawValue, out _))
                        return new ValidationError { RowNumber = rowIndex, ColumnName = columnName, ErrorType = "date", Message = rule.Message ?? $"Campo '{columnName}' deve ser uma data válida", RawValue = rawValue };
                    break;
                case "in_list":
                    if (!string.IsNullOrWhiteSpace(rawValue) && rule.AllowedValues != null && !rule.AllowedValues.Any(v => string.Equals(v, rawValue, StringComparison.OrdinalIgnoreCase)))
                        return new ValidationError { RowNumber = rowIndex, ColumnName = columnName, ErrorType = "in_list", Message = rule.Message ?? $"Campo '{columnName}' valor '{rawValue}' não está na lista permitida", RawValue = rawValue };
                    break;
                case "regex":
                    if (!string.IsNullOrWhiteSpace(rawValue) && !string.IsNullOrWhiteSpace(rule.Pattern) && !System.Text.RegularExpressions.Regex.IsMatch(rawValue, rule.Pattern))
                        return new ValidationError { RowNumber = rowIndex, ColumnName = columnName, ErrorType = "regex", Message = rule.Message ?? $"Campo '{columnName}' não corresponde ao padrão", RawValue = rawValue };
                    break;
            }
            return null;
        }

        #endregion

        #region Transform

        public string ApplyTransforms(string value, TransformConfig[] transforms)
        {
            if (transforms == null || transforms.Length == 0) return value;
            var result = value;
            foreach (var t in transforms)
            {
                switch (t.Type?.ToLower())
                {
                    case "trim": result = result.Trim(); break;
                    case "upper_case": result = result.ToUpper(); break;
                    case "lower_case": result = result.ToLower(); break;
                    case "title_case": if (!string.IsNullOrEmpty(result)) result = char.ToUpper(result[0]) + result.Substring(1).ToLower(); break;
                    case "replace": if (t.Params != null && t.Params.ContainsKey("from") && t.Params.ContainsKey("to")) result = result.Replace(t.Params["from"].ToString(), t.Params["to"].ToString()); break;
                    case "to_decimal": if (decimal.TryParse(result, out var dv)) result = dv.ToString(System.Globalization.CultureInfo.InvariantCulture); break;
                    case "to_int": if (int.TryParse(result, out var iv)) result = iv.ToString(); break;
                    case "to_date": if (!string.IsNullOrWhiteSpace(t.DateFormat) && DateTime.TryParseExact(result, t.DateFormat, null, System.Globalization.DateTimeStyles.None, out var dtv)) result = dtv.ToString("yyyy-MM-dd"); break;
                    case "to_boolean": if (t.TrueValues != null && t.TrueValues.Any(v => string.Equals(v, result, StringComparison.OrdinalIgnoreCase))) result = "true"; else if (t.FalseValues != null && t.FalseValues.Any(v => string.Equals(v, result, StringComparison.OrdinalIgnoreCase))) result = "false"; break;
                    case "default": if (string.IsNullOrWhiteSpace(result) && t.DefaultValue != null) result = t.DefaultValue.ToString(); break;
                    case "split": if (!string.IsNullOrWhiteSpace(t.Delimiter) && t.SplitIndex.HasValue) { var parts = result.Split(new[] { t.Delimiter }, StringSplitOptions.None); if (t.SplitIndex.Value < parts.Length) result = parts[t.SplitIndex.Value]; } break;
                    case "map_values": if (t.ValueMapping != null && t.ValueMapping.ContainsKey(result)) result = t.ValueMapping[result].ToString(); break;
                }
            }
            return result;
        }

        #endregion

        #region Export (data → Excel)

        public byte[] TabularToExcel(IEnumerable<IDictionary<string, object>> data, string sheetName = "Sheet1", bool hasHeaders = true)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add(sheetName);
                var rowIndex = 1;
                var dataList = data?.ToList();
                if (dataList == null || dataList.Count == 0)
                {
                    using (var stream = new MemoryStream()) { workbook.SaveAs(stream); return stream.ToArray(); }
                }

                var headers = dataList[0].Keys.ToList();
                if (hasHeaders)
                {
                    for (var col = 0; col < headers.Count; col++)
                    {
                        worksheet.Cell(rowIndex, col + 1).Value = headers[col];
                        worksheet.Cell(rowIndex, col + 1).Style.Font.Bold = true;
                    }
                    rowIndex++;
                }

                foreach (var row in dataList)
                {
                    for (var col = 0; col < headers.Count; col++)
                    {
                        var value = row.ContainsKey(headers[col]) ? row[headers[col]] : null;
                        worksheet.Cell(rowIndex, col + 1).Value = value?.ToString() ?? string.Empty;
                    }
                    rowIndex++;
                }

                worksheet.Columns().AdjustToContents();
                using (var stream = new MemoryStream()) { workbook.SaveAs(stream); return stream.ToArray(); }
            }
        }

        public byte[] CsvToExcel(Stream csvContent, string delimiter = ",", string sheetName = "Sheet1")
        {
            var rows = new List<IDictionary<string, object>>();
            string[] headers = null;
            using (var reader = new StreamReader(csvContent, Encoding.UTF8))
            {
                string line;
                var lineIndex = 0;
                while ((line = reader.ReadLine()) != null)
                {
                    var fields = ParseCsvLine(line, delimiter);
                    if (lineIndex == 0) { headers = fields; }
                    else
                    {
                        var row = new Dictionary<string, object>();
                        for (var i = 0; i < headers.Length; i++)
                            row[headers[i]] = i < fields.Length ? fields[i] : string.Empty;
                        rows.Add(row);
                    }
                    lineIndex++;
                }
            }
            return TabularToExcel(rows, sheetName);
        }

        public byte[] CsvToExcel(string csvContent, string delimiter = ",", string sheetName = "Sheet1")
        {
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(csvContent)))
            {
                return CsvToExcel(stream, delimiter, sheetName);
            }
        }

        private static string[] ParseCsvLine(string line, string delimiter)
        {
            var result = new List<string>();
            var current = new StringBuilder();
            var inQuotes = false;
            for (var i = 0; i < line.Length; i++)
            {
                var c = line[i];
                if (c == '"') { inQuotes = !inQuotes; }
                else if (!inQuotes && delimiter.Length == 1 && c == delimiter[0]) { result.Add(current.ToString().Trim()); current.Clear(); }
                else { current.Append(c); }
            }
            result.Add(current.ToString().Trim());
            return result.ToArray();
        }

        #endregion

        #region Reverse Flow (data ← JSON/CSV/XML)

        public byte[] JsonToExcel(string jsonContent, string sheetName = "Sheet1", bool hasHeaders = true)
        {
            if (string.IsNullOrWhiteSpace(jsonContent)) throw new Exception("JSON content not informed");
            var rows = ParseJsonToTabular(jsonContent);
            return TabularToExcel(rows, sheetName, hasHeaders);
        }

        public byte[] JsonToExcel(Stream jsonContent, string sheetName = "Sheet1", bool hasHeaders = true)
        {
            if (jsonContent == null) throw new Exception("JSON content not informed");
            using (var reader = new StreamReader(jsonContent, Encoding.UTF8))
            {
                return JsonToExcel(reader.ReadToEnd(), sheetName, hasHeaders);
            }
        }

        public byte[] JsonToCsv(string jsonContent, string delimiter = ",")
        {
            if (string.IsNullOrWhiteSpace(jsonContent)) throw new Exception("JSON content not informed");
            var rows = ParseJsonToTabular(jsonContent);
            if (rows.Count == 0) return Encoding.UTF8.GetBytes(string.Empty);

            var sb = new StringBuilder();
            var headers = rows[0].Keys.ToList();
            sb.AppendLine(string.Join(delimiter, headers.Select(h => EscapeCsvField(h, delimiter))));
            foreach (var row in rows)
            {
                var fields = headers.Select(h => EscapeCsvField(row.ContainsKey(h) ? row[h]?.ToString() ?? string.Empty : string.Empty, delimiter));
                sb.AppendLine(string.Join(delimiter, fields));
            }
            return Encoding.UTF8.GetBytes(sb.ToString());
        }

        public Stream JsonToCsvStream(string jsonContent, string delimiter = ",")
        {
            return new MemoryStream(JsonToCsv(jsonContent, delimiter));
        }

        public byte[] XmlToExcel(string xmlContent, string sheetName = "Sheet1", bool hasHeaders = true)
        {
            if (string.IsNullOrWhiteSpace(xmlContent)) throw new Exception("XML content not informed");
            var rows = ParseXmlToTabular(xmlContent);
            return TabularToExcel(rows, sheetName, hasHeaders);
        }

        public byte[] XmlToExcel(Stream xmlContent, string sheetName = "Sheet1", bool hasHeaders = true)
        {
            if (xmlContent == null) throw new Exception("XML content not informed");
            using (var reader = new StreamReader(xmlContent, Encoding.UTF8))
            {
                return XmlToExcel(reader.ReadToEnd(), sheetName, hasHeaders);
            }
        }

        public byte[] XmlToCsv(string xmlContent, string delimiter = ",")
        {
            if (string.IsNullOrWhiteSpace(xmlContent)) throw new Exception("XML content not informed");
            var rows = ParseXmlToTabular(xmlContent);
            if (rows.Count == 0) return Encoding.UTF8.GetBytes(string.Empty);

            var sb = new StringBuilder();
            var headers = rows[0].Keys.ToList();
            sb.AppendLine(string.Join(delimiter, headers.Select(h => EscapeCsvField(h, delimiter))));
            foreach (var row in rows)
            {
                var fields = headers.Select(h => EscapeCsvField(row.ContainsKey(h) ? row[h]?.ToString() ?? string.Empty : string.Empty, delimiter));
                sb.AppendLine(string.Join(delimiter, fields));
            }
            return Encoding.UTF8.GetBytes(sb.ToString());
        }

        #endregion

        #region Form Sheet Parser

        public string GetJsonStringFromForm(string fileName, string sheetName, string[] replaceFrom = null, string[] replaceTo = null, string[] fieldNames = null)
        {
            if (string.IsNullOrWhiteSpace(fileName)) throw new Exception("File name not informed");
            using (var engine = new XLWorkbook(fileName))
            {
                var parsedData = ParseFormSheet(engine, sheetName, fieldNames);
                return WriteJsonBodyFromNamedFields(parsedData);
            }
        }

        public string GetJsonStringFromForm(Stream fileContent, string sheetName, string[] replaceFrom = null, string[] replaceTo = null, string[] fieldNames = null)
        {
            if (fileContent == null) throw new Exception("File content not informed");
            using (var engine = new XLWorkbook(fileContent))
            {
                var parsedData = ParseFormSheet(engine, sheetName, fieldNames);
                return WriteJsonBodyFromNamedFields(parsedData);
            }
        }

        public object GetJsonObjectFromForm(string fileName, string sheetName, string[] replaceFrom = null, string[] replaceTo = null, string[] fieldNames = null)
        {
            var strJson = GetJsonStringFromForm(fileName, sheetName, replaceFrom, replaceTo, fieldNames);
            return JsonSerializer.Deserialize<object>(strJson);
        }

        public object GetJsonObjectFromForm(Stream fileContent, string sheetName, string[] replaceFrom = null, string[] replaceTo = null, string[] fieldNames = null)
        {
            var strJson = GetJsonStringFromForm(fileContent, sheetName, replaceFrom, replaceTo, fieldNames);
            return JsonSerializer.Deserialize<object>(strJson);
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
                var className = Path.GetFileNameWithoutExtension(fileName);
                result = generator.GenerateFile(className);
            }
            return result;
        }

        public IDictionary<string, object> GetDictionary(string fileName, string sheetName, string[] replaceFrom = null, string[] replaceTo = null, string[] fieldNames = null)
        {
            if (string.IsNullOrWhiteSpace(fileName)) throw new Exception("File name not informed");
            using (var engine = new XLWorkbook(fileName))
            {
                return ParseFormSheet(engine, sheetName, fieldNames);
            }
        }

        public IDictionary<string, object> GetDictionary(Stream fileContent, string sheetName, string[] replaceFrom = null, string[] replaceTo = null, string[] fieldNames = null)
        {
            if (fileContent == null) throw new Exception("File content not informed");
            using (var engine = new XLWorkbook(fileContent))
            {
                return ParseFormSheet(engine, sheetName, fieldNames);
            }
        }

        #endregion

        #region Helpers

        private static Stream GetFileStream(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName)) throw new Exception("File name not informed");
            return File.Open(fileName, FileMode.Open, FileAccess.Read);
        }

        private static string[] GetHeaderColumns(IExcelDataReader reader)
        {
            if (reader == null) return null;
            var result = new string[reader.FieldCount];
            for (var count = 0; count < reader.FieldCount; count++)
                result[count] = reader[count]?.ToString().Trim() ?? $"Column{count + 1}";
            return result;
        }

        private static void ApplyColumnNamesReplace(string[] columnNames, string[] readFrom, string[] replaceTo)
        {
            if (readFrom == null || replaceTo == null) return;
            if (readFrom.Length != replaceTo.Length) throw new ArgumentOutOfRangeException("Invalid replace values amount");
            for (var nameCount = 0; nameCount < columnNames.Length; nameCount++)
                for (var chrCount = 0; chrCount < readFrom.Length; chrCount++)
                    columnNames[nameCount] = columnNames[nameCount].Replace(readFrom[chrCount], replaceTo[chrCount]).Replace(" ", "_");
        }

        private static void WriteItemJsonBodyFromReader(IExcelDataReader reader, Utf8JsonWriter writer, string[] headerColumns)
        {
            writer.WriteStartObject();
            var colCount = 0;
            foreach (var col in headerColumns)
            {
                var colValue = reader.GetValue(colCount);
                writer.WritePropertyName(col);
                if (colValue == null) writer.WriteNullValue();
                else if (colValue is int intVal) writer.WriteNumberValue(intVal);
                else if (colValue is double doubleVal) writer.WriteNumberValue(doubleVal);
                else if (colValue is decimal decimalVal) writer.WriteNumberValue((double)decimalVal);
                else if (colValue is bool boolVal) writer.WriteBooleanValue(boolVal);
                else if (colValue is DateTime dateVal) writer.WriteStringValue(dateVal.ToString("yyyy-MM-dd HH:mm:ss"));
                else writer.WriteStringValue(colValue.ToString());
                colCount++;
            }
            writer.WriteEndObject();
        }

        private static List<IDictionary<string, object>> ParseJsonToTabular(string jsonContent)
        {
            var rows = new List<IDictionary<string, object>>();
            using var doc = JsonDocument.Parse(jsonContent);
            var root = doc.RootElement;

            if (root.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in root.EnumerateArray())
                {
                    var row = new Dictionary<string, object>();
                    if (item.ValueKind == JsonValueKind.Object)
                        foreach (var prop in item.EnumerateObject())
                            row[prop.Name] = JsonElementToObject(prop.Value);
                    rows.Add(row);
                }
            }
            else if (root.ValueKind == JsonValueKind.Object)
            {
                var row = new Dictionary<string, object>();
                foreach (var prop in root.EnumerateObject())
                    row[prop.Name] = JsonElementToObject(prop.Value);
                rows.Add(row);
            }
            return rows;
        }

        private static object JsonElementToObject(JsonElement element)
        {
            switch (element.ValueKind)
            {
                case JsonValueKind.String: return element.GetString();
                case JsonValueKind.Number:
                    if (element.TryGetInt32(out var intVal)) return intVal;
                    if (element.TryGetInt64(out var longVal)) return longVal;
                    return element.GetDouble();
                case JsonValueKind.True: return true;
                case JsonValueKind.False: return false;
                case JsonValueKind.Null: return null;
                case JsonValueKind.Array:
                    var list = new List<object>();
                    foreach (var item in element.EnumerateArray()) list.Add(JsonElementToObject(item));
                    return list;
                case JsonValueKind.Object:
                    var dict = new Dictionary<string, object>();
                    foreach (var prop in element.EnumerateObject()) dict[prop.Name] = JsonElementToObject(prop.Value);
                    return dict;
                default: return element.ToString();
            }
        }

        private static List<IDictionary<string, object>> ParseXmlToTabular(string xmlContent)
        {
            var rows = new List<IDictionary<string, object>>();
            var doc = XDocument.Parse(xmlContent);
            var root = doc.Root;
            if (root == null) return rows;

            var records = root.Elements();
            if (!records.Any()) return rows;

            var headers = records.First().Elements().Select(e => e.Name.LocalName).Distinct().ToList();
            foreach (var record in records)
            {
                var row = new Dictionary<string, object>();
                foreach (var header in headers)
                {
                    var element = record.Element(header);
                    row[header] = element?.Value ?? string.Empty;
                }
                rows.Add(row);
            }
            return rows;
        }

        private static string EscapeCsvField(string field, string delimiter)
        {
            if (string.IsNullOrEmpty(field)) return string.Empty;
            if (field.Contains(delimiter) || field.Contains('"') || field.Contains('\n'))
                return "\"" + field.Replace("\"", "\"\"") + "\"";
            return field;
        }

        private static IDictionary<string, object> ParseFormSheet(XLWorkbook engine, string sheetName, string[] fieldNames = null)
        {
            if (fieldNames == null)
                fieldNames = engine.NamedRanges.Select(nmf => nmf.Name).ToArray();

            var result = new Dictionary<string, object>();
            foreach (var field in fieldNames)
            {
                var cell = engine.Cell(field);
                if (cell != null && cell.Worksheet.Name.ToLower().Equals(sheetName.ToLower()))
                {
                    try { result.Add(field, cell.Value); }
                    catch (InvalidOperationException) { throw new InvalidOperationException($"Invalid cell value at {field}."); }
                }
            }
            return result;
        }

        private static string WriteJsonBodyFromNamedFields(IDictionary<string, object> fields)
        {
            using (var ms = new MemoryStream())
            {
                using (var writer = new Utf8JsonWriter(ms, new JsonWriterOptions { Indented = true }))
                {
                    if (fields != null)
                    {
                        writer.WriteStartObject();
                        foreach (var field in fields)
                        {
                            writer.WritePropertyName(field.Key);
                            if (field.Value != null)
                            {
                                if (field.Value is int iv) writer.WriteNumberValue(iv);
                                else if (field.Value is double dv) writer.WriteNumberValue(dv);
                                else if (field.Value is decimal dcv) writer.WriteNumberValue((double)dcv);
                                else if (field.Value is bool bv) writer.WriteBooleanValue(bv);
                                else writer.WriteStringValue(field.Value.ToString());
                            }
                            else writer.WriteNullValue();
                        }
                        writer.WriteEndObject();
                    }
                }
                return Encoding.UTF8.GetString(ms.ToArray());
            }
        }

        #endregion

        #region IDisposable

        private bool _disposed;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing) { }
                _disposed = true;
            }
        }

        #endregion
    }

    #region Support Models

    public class ValidationResult
    {
        public bool IsValid { get; set; }
        public int TotalRows { get; set; }
        public int ValidRows { get; set; }
        public int ErrorRows { get; set; }
        public List<ValidationError> Errors { get; set; } = new List<ValidationError>();
    }

    public class ValidationError
    {
        public int RowNumber { get; set; }
        public string ColumnName { get; set; } = string.Empty;
        public string ErrorType { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
        public string RawValue { get; set; } = string.Empty;
    }

    public class ValidationRule
    {
        public string ColumnName { get; set; } = string.Empty;
        public string Type { get; set; } = string.Empty;
        public string Message { get; set; }
        public int? MaxLength { get; set; }
        public int? MinLength { get; set; }
        public string Pattern { get; set; }
        public string[] AllowedValues { get; set; }
    }

    public class TransformConfig
    {
        public string Type { get; set; } = string.Empty;
        public Dictionary<string, object> Params { get; set; }
        public string DateFormat { get; set; }
        public string[] TrueValues { get; set; }
        public string[] FalseValues { get; set; }
        public object DefaultValue { get; set; }
        public string Delimiter { get; set; }
        public int? SplitIndex { get; set; }
        public Dictionary<string, object> ValueMapping { get; set; }
    }

    #endregion
}
