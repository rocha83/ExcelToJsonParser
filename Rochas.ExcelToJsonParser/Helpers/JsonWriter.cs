using System;
using System.IO;
using Newtonsoft.Json;
using ExcelDataReader;

namespace Rochas.ExcelToJson.Helpers
{
    internal static class JsonWriterHelper
    {
        public static void WriteItemJsonBodyFromReader(IExcelDataReader reader, JsonWriter writer, string[] headerColumns)
        {
            writer.WriteStartObject();

            var colCount = 0;
            foreach (var col in headerColumns)
            {
                var colValue = reader.GetValue(colCount);
                writer.WritePropertyName(col);
                writer.WriteValue(colValue);
                colCount += 1;
            }

            writer.WriteEndObject();
        }

        public static string WriteJsonBodyFromNamedFields(System.Collections.Generic.IDictionary<string, object> fields)
        {
            using (var result = new StringWriter())
            using (var writer = new JsonTextWriter(result))
            {
                writer.Formatting = Formatting.Indented;

                if ((fields != null) && (writer != null))
                {
                    writer.WriteStartObject();

                    foreach (var field in fields)
                    {
                        writer.WritePropertyName(field.Key);
                        if (field.Value != null)
                            writer.WriteValue(field.Value);
                    }

                    writer.WriteEndObject();
                }

                return result.ToString();
            }
        }
    }
}
