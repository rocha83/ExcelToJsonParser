using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using ExcelDataReader;

namespace Rochas.ExcelToJson.Helpers
{
    internal static class TabularSheetReader
    {
        public static Stream GetFileStream(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                throw new Exception("File name not informed");

            return File.Open(fileName, FileMode.Open, FileAccess.Read);
        }

        public static string[] GetHeaderColumns(IExcelDataReader reader)
        {
            string[] result = null;

            if (reader != null)
            {
                result = new string[reader.FieldCount];

                for (var count = 0; count < reader.FieldCount; count++)
                    result[count] = reader[count]?.ToString().Trim();
            }

            return result;
        }

        public static DataTable GetDataTable(Stream fileContent, int skipRows, bool useHeader)
        {
            DataTable result = null;
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            if (fileContent != null)
            {
                var reader = ExcelReaderFactory.CreateReader(fileContent);

                var config = new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = useHeader
                    }
                };

                while (skipRows > 0)
                {
                    reader.Read();
                    skipRows--;
                }

                result = reader.AsDataSet(config).Tables[0];
            }

            return result;
        }
    }
}
