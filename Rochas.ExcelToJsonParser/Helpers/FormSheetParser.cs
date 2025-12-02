using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace Rochas.ExcelToJson.Helpers
{
    internal static class FormSheetHelper
    {
        public static string[] GetNamedFormFields(XLWorkbook engine)
        {
            if (engine != null)
                return engine.DefinedNames.Select(nmf => nmf.Name).ToArray();

            return null;
        }

        public static IDictionary<string, object> GetNamedFieldValues(XLWorkbook engine, string sheetName, string[] fieldNames)
        {
            IDictionary<string, object> result = new Dictionary<string, object>();

            foreach (var field in fieldNames)
            {
                var cell = engine.Cell(field);

                if (cell != null)
                    if (cell.Worksheet.Name.ToLower().Equals(sheetName.ToLower()))
                    {
                        try
                        {
                            result.Add(field, cell.Value);
                        }
                        catch (Exception ex)
                        {
                            if (ex.GetType() == typeof(InvalidOperationException))
                                throw new InvalidOperationException($"Invalid cell value at {field}.");
                        }
                    }
            }

            return result;
        }

        public static IDictionary<string, object> ParseFormSheet(XLWorkbook engine, string sheetName, string[] fieldNames)
        {
            if (fieldNames == null)
                fieldNames = GetNamedFormFields(engine);

            return GetNamedFieldValues(engine, sheetName, fieldNames);
        }
    }
}
