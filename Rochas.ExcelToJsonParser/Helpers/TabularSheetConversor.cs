using System;
using ExcelDataReader;

namespace Rochas.ExcelToJson.Helpers
{
    internal static class TabularSheetConversor
    {
        public static void ApplyColumnNamesReplace(string[] columnNames, string[] readFrom, string[] replaceTo)
        {
            if ((readFrom != null) && (replaceTo != null))
            {
                if (readFrom.Length != replaceTo.Length)
                    throw new ArgumentOutOfRangeException("Invalid replace values amount");

                for (var nameCount = 0; nameCount < columnNames.Length; nameCount++)
                {
                    for (var chrCount = 0; chrCount < readFrom.Length; chrCount++)
                        columnNames[nameCount] = columnNames[nameCount]
                            .Replace(readFrom[chrCount], replaceTo[chrCount])
                            .Replace(" ", "_");
                }
            }
        }
    }
}
