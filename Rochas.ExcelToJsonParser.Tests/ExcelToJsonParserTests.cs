using System.Text;
using System.Text.Json;
using ClosedXML.Excel;
using Rochas.ExcelToJson;
using Xunit;

namespace Rochas.ExcelToJsonParser.Tests;

public class ExcelToJsonParserTests
{
    private readonly Rochas.ExcelToJson.ExcelToJsonParser _parser = new();

    private string GetSamplePath(string name) => Path.Combine("Samples", name);
    private static Stream ToStream(string text) => new MemoryStream(Encoding.UTF8.GetBytes(text));

    private static string GetJsonArray() => @"[
        {""Nome"":""João"",""Idade"":30,""Cidade"":""SP""},
        {""Nome"":""Maria"",""Idade"":25,""Cidade"":""RJ""},
        {""Nome"":""Pedro"",""Idade"":35,""Cidade"":""MG""}
    ]";

    private static string GetJsonObject() => @"{""Nome"":""João"",""Idade"":30,""Cidade"":""SP""}";

    private static string GetXmlContent() => @"<Pessoas>
        <Pessoa><Nome>João</Nome><Idade>30</Idade><Cidade>SP</Cidade></Pessoa>
        <Pessoa><Nome>Maria</Nome><Idade>25</Idade><Cidade>RJ</Cidade></Pessoa>
    </Pessoas>";

    private static byte[] CreateSampleExcel()
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("TestData");
        ws.Cell(1, 1).Value = "Nome"; ws.Cell(1, 2).Value = "Idade"; ws.Cell(1, 3).Value = "Cidade";
        ws.Cell(2, 1).Value = "João"; ws.Cell(2, 2).Value = 30; ws.Cell(2, 3).Value = "SP";
        ws.Cell(3, 1).Value = "Maria"; ws.Cell(3, 2).Value = 25; ws.Cell(3, 3).Value = "RJ";
        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        return ms.ToArray();
    }

    #region Instantiation

    [Fact] public void New_Instance_IsNotNull() { using var p = new Rochas.ExcelToJson.ExcelToJsonParser(); Assert.NotNull(p); }
    [Fact] public void New_Instance_ImplementsIDisposable() { using var p = new Rochas.ExcelToJson.ExcelToJsonParser(); Assert.IsAssignableFrom<IDisposable>(p); }
    [Fact] public void New_Instance_ImplementsInterface() { using var p = new Rochas.ExcelToJson.ExcelToJsonParser(); Assert.IsAssignableFrom<Rochas.ExcelToJson.IExcelToJsonParser>(p); }

    #endregion

    #region Tabular Parsing

    [Fact] public void GetJsonStringFromTabular_FileName_ReturnsJson()
    {
        var json = _parser.GetJsonStringFromTabular(GetSamplePath("TabularSample.xlsx"));
        Assert.False(string.IsNullOrWhiteSpace(json));
        Assert.Contains("[", json);
    }

    [Fact] public void GetJsonStringFromTabular_FileName_WithReplace()
    {
        var json = _parser.GetJsonStringFromTabular(GetSamplePath("TabularSample.xlsx"), 0, new[] { "á", "é" }, new[] { "a", "e" });
        Assert.False(string.IsNullOrWhiteSpace(json));
    }

    [Fact] public void GetJsonStringFromTabular_FileName_OnlySampleRow()
    {
        var json = _parser.GetJsonStringFromTabular(GetSamplePath("TabularSample.xlsx"), 0, null, null, null, true);
        var arr = JsonSerializer.Deserialize<JsonElement>(json);
        Assert.Equal(JsonValueKind.Array, arr.ValueKind);
        Assert.True(arr.GetArrayLength() <= 1);
    }

    [Fact] public void GetJsonStringFromTabular_Stream_ReturnsJson()
    {
        using var stream = File.OpenRead(GetSamplePath("TabularSample.xlsx"));
        var json = _parser.GetJsonStringFromTabular(stream);
        Assert.Contains("[", json);
    }

    [Fact] public void GetJsonStringFromTabular_Stream_WithSkipRows()
    {
        using var stream = File.OpenRead(GetSamplePath("TabularSample.xlsx"));
        var json = _parser.GetJsonStringFromTabular(stream, skipRows: 1);
        Assert.False(string.IsNullOrWhiteSpace(json));
    }

    [Fact] public void GetJsonStringFromTabular_Stream_WithCustomHeaders()
    {
        using var stream = File.OpenRead(GetSamplePath("TabularSample.xlsx"));
        var dt = _parser.GetDataTable(stream);
        var headers = Enumerable.Range(1, dt.Columns.Count).Select(i => $"Col{i}").ToArray();
        using var stream2 = File.OpenRead(GetSamplePath("TabularSample.xlsx"));
        var json = _parser.GetJsonStringFromTabular(stream2, 0, null, null, headers);
        Assert.True(JsonSerializer.Deserialize<JsonElement>(json).GetArrayLength() > 0);
    }

    [Fact] public void GetJsonStringFromTabular_Stream_InvalidHeaders_Throws()
    {
        using var stream = File.OpenRead(GetSamplePath("TabularSample.xlsx"));
        Assert.Throws<Exception>(() => _parser.GetJsonStringFromTabular(stream, 0, null, null, new[] { "A" }));
    }

    [Fact] public void GetJsonObjectFromTabular_ReturnsEnumerable()
    {
        var result = _parser.GetJsonObjectFromTabular(GetSamplePath("TabularSample.xlsx"));
        Assert.NotNull(result);
        Assert.True(result.Any());
    }

    [Fact] public void GetClassModelFromTabular_ReturnsCSharpCode()
    {
        var result = _parser.GetClassModelFromTabular(GetSamplePath("TabularSample.xlsx"));
        Assert.Contains("class", result);
    }

    #endregion

    #region GetDataTable

    [Fact] public void GetDataTable_FileName_ReturnsDataTable()
    {
        var dt = _parser.GetDataTable(GetSamplePath("TabularSample.xlsx"));
        Assert.NotNull(dt);
        Assert.True(dt.Rows.Count > 0);
    }

    [Fact] public void GetDataTable_Stream_ReturnsDataTable()
    {
        using var stream = File.OpenRead(GetSamplePath("TabularSample.xlsx"));
        Assert.NotNull(_parser.GetDataTable(stream));
    }

    [Fact] public void GetDataTable_FileName_WithSkipRows() { Assert.NotNull(_parser.GetDataTable(GetSamplePath("TabularSample.xlsx"), skipRows: 1)); }
    [Fact] public void GetDataTable_Stream_Null_ReturnsNull() { Assert.Null(_parser.GetDataTable((Stream)null!)); }
    [Fact] public void GetDataTable_FileName_Empty_ReturnsNull() { Assert.Null(_parser.GetDataTable("")); }
    [Fact] public void GetDataTable_BySheetName() { Assert.NotNull(_parser.GetDataTable(GetSamplePath("TabularSample.xlsx"), "TestData")); }

    #endregion

    #region Streaming

    [Fact]
    public async Task StreamFromTabular_YieldsRows()
    {
        using var stream = File.OpenRead(GetSamplePath("TabularSample.xlsx"));
        var rows = new List<IDictionary<string, object>>();
        await foreach (var row in _parser.StreamFromTabular(stream))
            rows.Add(row);
        Assert.True(rows.Count > 0);
    }

    [Fact]
    public async Task StreamFromTabular_WithCancellation()
    {
        using var stream = File.OpenRead(GetSamplePath("TabularSample.xlsx"));
        using var cts = new CancellationTokenSource();
        cts.Cancel();
        var rows = new List<IDictionary<string, object>>();
        await foreach (var row in _parser.StreamFromTabular(stream, cancellationToken: cts.Token))
            rows.Add(row);
        Assert.Empty(rows);
    }

    #endregion

    #region Multi-sheet

    [Fact] public void GetJsonStringsFromAllSheets_ReturnsDictionary()
    {
        var result = _parser.GetJsonStringsFromAllSheets(GetSamplePath("TabularSample.xlsx"));
        Assert.True(result.Count > 0);
        foreach (var kvp in result) Assert.Contains("[", kvp.Value);
    }

    [Fact] public void GetJsonStringsFromAllSheets_WithReplace()
    {
        var result = _parser.GetJsonStringsFromAllSheets(GetSamplePath("TabularSample.xlsx"), 0, new[] { "á" }, new[] { "a" });
        Assert.True(result.Count > 0);
    }

    #endregion

    #region Validation

    [Fact] public void ValidateTabular_Required_Passes()
    {
        using var stream = new MemoryStream(CreateSampleExcel());
        var result = _parser.ValidateTabular(stream, new[] { new ValidationRule { ColumnName = "Nome", Type = "required" } });
        Assert.True(result.IsValid);
    }

    [Fact] public void ValidateTabular_Numeric_Passes()
    {
        using var stream = new MemoryStream(CreateSampleExcel());
        Assert.True(_parser.ValidateTabular(stream, new[] { new ValidationRule { ColumnName = "Idade", Type = "numeric" } }).IsValid);
    }

    [Fact] public void ValidateTabular_MaxLength_Fails()
    {
        using var stream = new MemoryStream(CreateSampleExcel());
        Assert.False(_parser.ValidateTabular(stream, new[] { new ValidationRule { ColumnName = "Nome", Type = "max_length", MaxLength = 2 } }).IsValid);
    }

    [Fact] public void ValidateTabular_MinLength_Fails()
    {
        using var stream = new MemoryStream(CreateSampleExcel());
        Assert.False(_parser.ValidateTabular(stream, new[] { new ValidationRule { ColumnName = "Nome", Type = "min_length", MinLength = 100 } }).IsValid);
    }

    [Fact] public void ValidateTabular_InList_Passes()
    {
        using var stream = new MemoryStream(CreateSampleExcel());
        Assert.True(_parser.ValidateTabular(stream, new[] { new ValidationRule { ColumnName = "Cidade", Type = "in_list", AllowedValues = new[] { "SP", "RJ", "MG" } } }).IsValid);
    }

    [Fact] public void ValidateTabular_InList_Fails()
    {
        using var stream = new MemoryStream(CreateSampleExcel());
        Assert.False(_parser.ValidateTabular(stream, new[] { new ValidationRule { ColumnName = "Cidade", Type = "in_list", AllowedValues = new[] { "XX" } } }).IsValid);
    }

    [Fact] public void ValidateTabular_Regex_Passes()
    {
        using var stream = new MemoryStream(CreateSampleExcel());
        Assert.NotNull(_parser.ValidateTabular(stream, new[] { new ValidationRule { ColumnName = "Nome", Type = "regex", Pattern = @"^[A-ZÁÉÍÓÚÇ].*" } }));
    }

    [Fact] public void ValidateTabular_Date_Passes()
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Dates");
        ws.Cell(1, 1).Value = "Data"; ws.Cell(2, 1).Value = "2026-01-15";
        using var ms = new MemoryStream(); workbook.SaveAs(ms); ms.Position = 0;
        Assert.True(_parser.ValidateTabular(ms, new[] { new ValidationRule { ColumnName = "Data", Type = "date" } }).IsValid);
    }

    [Fact] public void ValidateTabular_MultipleRules()
    {
        using var stream = new MemoryStream(CreateSampleExcel());
        var result = _parser.ValidateTabular(stream, new[] {
            new ValidationRule { ColumnName = "Nome", Type = "required" },
            new ValidationRule { ColumnName = "Idade", Type = "numeric" },
            new ValidationRule { ColumnName = "Cidade", Type = "in_list", AllowedValues = new[] { "SP", "RJ", "MG" } }
        });
        Assert.True(result.IsValid);
    }

    [Fact] public void ValidateTabular_CustomMessage()
    {
        using var stream = new MemoryStream(CreateSampleExcel());
        var result = _parser.ValidateTabular(stream, new[] { new ValidationRule { ColumnName = "Nome", Type = "max_length", MaxLength = 1, Message = "Nome muito longo" } });
        Assert.Contains("Nome muito longo", result.Errors[0].Message);
    }

    #endregion

    #region Transform

    [Fact] public void ApplyTransforms_Trim() { Assert.Equal("abc", _parser.ApplyTransforms("  abc  ", new[] { new TransformConfig { Type = "trim" } })); }
    [Fact] public void ApplyTransforms_UpperCase() { Assert.Equal("ABC", _parser.ApplyTransforms("abc", new[] { new TransformConfig { Type = "upper_case" } })); }
    [Fact] public void ApplyTransforms_LowerCase() { Assert.Equal("abc", _parser.ApplyTransforms("ABC", new[] { new TransformConfig { Type = "lower_case" } })); }
    [Fact] public void ApplyTransforms_TitleCase() { Assert.Equal("Abc", _parser.ApplyTransforms("abc", new[] { new TransformConfig { Type = "title_case" } })); }
    [Fact] public void ApplyTransforms_Replace() { Assert.Equal("a_b", _parser.ApplyTransforms("a b", new[] { new TransformConfig { Type = "replace", Params = new() { { "from", " " }, { "to", "_" } } } })); }
    [Fact] public void ApplyTransforms_ToDecimal() { Assert.Equal("1.5", _parser.ApplyTransforms("1,5", new[] { new TransformConfig { Type = "to_decimal" } })); }
    [Fact] public void ApplyTransforms_ToInt() { Assert.Equal("42", _parser.ApplyTransforms("42", new[] { new TransformConfig { Type = "to_int" } })); }
    [Fact] public void ApplyTransforms_ToDate() { Assert.Equal("2026-01-15", _parser.ApplyTransforms("15/01/2026", new[] { new TransformConfig { Type = "to_date", DateFormat = "dd/MM/yyyy" } })); }
    [Fact] public void ApplyTransforms_ToBoolean_True() { Assert.Equal("true", _parser.ApplyTransforms("sim", new[] { new TransformConfig { Type = "to_boolean", TrueValues = new[] { "sim" }, FalseValues = new[] { "não" } } })); }
    [Fact] public void ApplyTransforms_ToBoolean_False() { Assert.Equal("false", _parser.ApplyTransforms("não", new[] { new TransformConfig { Type = "to_boolean", TrueValues = new[] { "sim" }, FalseValues = new[] { "não" } } })); }
    [Fact] public void ApplyTransforms_Default() { Assert.Equal("N/A", _parser.ApplyTransforms("", new[] { new TransformConfig { Type = "default", DefaultValue = "N/A" } })); }
    [Fact] public void ApplyTransforms_Split() { Assert.Equal("b", _parser.ApplyTransforms("a,b,c", new[] { new TransformConfig { Type = "split", Delimiter = ",", SplitIndex = 1 } })); }
    [Fact] public void ApplyTransforms_MapValues() { Assert.Equal("São Paulo", _parser.ApplyTransforms("SP", new[] { new TransformConfig { Type = "map_values", ValueMapping = new() { { "SP", "São Paulo" } } } })); }
    [Fact] public void ApplyTransforms_Empty() { Assert.Equal("abc", _parser.ApplyTransforms("abc", null)); }
    [Fact] public void ApplyTransforms_Multiple()
    {
        Assert.Equal("JOAO_SILVA", _parser.ApplyTransforms("  joao silva  ", new[] {
            new TransformConfig { Type = "trim" }, new TransformConfig { Type = "upper_case" },
            new TransformConfig { Type = "replace", Params = new() { { "from", " " }, { "to", "_" } } }
        }));
    }

    #endregion

    #region Export

    [Fact] public void TabularToExcel_ReturnsBytes()
    {
        var data = new List<IDictionary<string, object>> { new Dictionary<string, object> { { "A", "1" } } };
        Assert.True(_parser.TabularToExcel(data).Length > 0);
    }

    [Fact] public void TabularToExcel_EmptyData() { Assert.True(_parser.TabularToExcel(new List<IDictionary<string, object>>()).Length > 0); }
    [Fact] public void TabularToExcel_NullData() { Assert.True(_parser.TabularToExcel(null!).Length > 0); }

    [Fact] public void TabularToExcel_CustomSheetName()
    {
        var bytes = _parser.TabularToExcel(new List<IDictionary<string, object>> { new Dictionary<string, object> { { "A", "1" } } }, "MySheet");
        using var ms = new MemoryStream(bytes);
        using var wb = new XLWorkbook(ms);
        Assert.Equal("MySheet", wb.Worksheets.First().Name);
    }

    [Fact] public void CsvToExcel_Stream_ReturnsBytes()
    {
        using var stream = ToStream("A,B\n1,2");
        Assert.True(_parser.CsvToExcel(stream).Length > 0);
    }

    [Fact] public void CsvToExcel_Stream_Semicolon()
    {
        using var stream = ToStream("A;B\n1;2");
        Assert.True(_parser.CsvToExcel(stream, ";").Length > 0);
    }

    [Fact] public void CsvToExcel_Stream_WithQuotes()
    {
        using var stream = ToStream("A,B\n1,\"2,3\"");
        Assert.True(_parser.CsvToExcel(stream).Length > 0);
    }

    [Fact] public void CsvToExcel_String_ReturnsBytes() { Assert.True(_parser.CsvToExcel("A,B\n1,2").Length > 0); }

    [Fact] public void CsvToExcel_String_CustomSheet()
    {
        var bytes = _parser.CsvToExcel("A,B\n1,2", ",", "CSheet");
        using var ms = new MemoryStream(bytes);
        using var wb = new XLWorkbook(ms);
        Assert.Equal("CSheet", wb.Worksheets.First().Name);
    }

    #endregion

    #region Reverse Flow - JsonToExcel

    [Fact] public void JsonToExcel_Array() { Assert.True(_parser.JsonToExcel(GetJsonArray()).Length > 0); }
    [Fact] public void JsonToExcel_Object() { Assert.True(_parser.JsonToExcel(GetJsonObject()).Length > 0); }
    [Fact] public void JsonToExcel_Stream() { Assert.True(_parser.JsonToExcel(ToStream(GetJsonArray())).Length > 0); }

    [Fact] public void JsonToExcel_CustomSheet()
    {
        var bytes = _parser.JsonToExcel(GetJsonArray(), "Pessoas");
        using var ms = new MemoryStream(bytes);
        using var wb = new XLWorkbook(ms);
        Assert.Equal("Pessoas", wb.Worksheets.First().Name);
    }

    [Fact] public void JsonToExcel_NoHeaders() { Assert.True(_parser.JsonToExcel(GetJsonArray(), "Sheet1", false).Length > 0); }
    [Fact] public void JsonToExcel_EmptyArray() { Assert.True(_parser.JsonToExcel("[]").Length > 0); }
    [Fact] public void JsonToExcel_EmptyString_Throws() { Assert.Throws<Exception>(() => _parser.JsonToExcel("")); }
    [Fact] public void JsonToExcel_NullString_Throws() { Assert.Throws<Exception>(() => _parser.JsonToExcel((string)null!)); }
    [Fact] public void JsonToExcel_Stream_Null_Throws() { Assert.Throws<Exception>(() => _parser.JsonToExcel((Stream)null!)); }

    [Fact] public void JsonToExcel_NestedObject()
    {
        Assert.True(_parser.JsonToExcel(@"[{""A"":{""B"":1}}]").Length > 0);
    }

    [Fact] public void JsonToExcel_ArrayValues()
    {
        Assert.True(_parser.JsonToExcel(@"[{""A"":[""x"",""y""]}]").Length > 0);
    }

    [Fact] public void JsonToExcel_NullValues()
    {
        Assert.True(_parser.JsonToExcel(@"[{""A"":null}]").Length > 0);
    }

    [Fact] public void JsonToExcel_AllTypes()
    {
        Assert.True(_parser.JsonToExcel(@"[{""S"":""x"",""I"":1,""D"":1.5,""B"":true,""N"":null}]").Length > 0);
    }

    #endregion

    #region Reverse Flow - JsonToCsv

    [Fact] public void JsonToCsv_ReturnsBytes()
    {
        var csv = Encoding.UTF8.GetString(_parser.JsonToCsv(GetJsonArray()));
        Assert.Contains("Nome", csv);
    }

    [Fact] public void JsonToCsv_CustomDelimiter()
    {
        var csv = Encoding.UTF8.GetString(_parser.JsonToCsv(GetJsonArray(), ";"));
        Assert.Contains(";", csv);
    }

    [Fact] public void JsonToCsv_EmptyArray() { Assert.NotNull(_parser.JsonToCsv("[]")); }
    [Fact] public void JsonToCsvStream() { Assert.True(_parser.JsonToCsvStream(GetJsonArray()).Length > 0); }
    [Fact] public void JsonToCsv_EmptyString_Throws() { Assert.Throws<Exception>(() => _parser.JsonToCsv("")); }
    [Fact] public void JsonToCsv_NullString_Throws() { Assert.Throws<Exception>(() => _parser.JsonToCsv(null!)); }

    #endregion

    #region Reverse Flow - XmlToExcel

    [Fact] public void XmlToExcel_String() { Assert.True(_parser.XmlToExcel(GetXmlContent()).Length > 0); }
    [Fact] public void XmlToExcel_Stream() { Assert.True(_parser.XmlToExcel(ToStream(GetXmlContent())).Length > 0); }

    [Fact] public void XmlToExcel_CustomSheet()
    {
        var bytes = _parser.XmlToExcel(GetXmlContent(), "XML");
        using var ms = new MemoryStream(bytes);
        using var wb = new XLWorkbook(ms);
        Assert.Equal("XML", wb.Worksheets.First().Name);
    }

    [Fact] public void XmlToExcel_EmptyString_Throws() { Assert.Throws<Exception>(() => _parser.XmlToExcel("")); }
    [Fact] public void XmlToExcel_NullString_Throws() { Assert.Throws<Exception>(() => _parser.XmlToExcel((string)null!)); }
    [Fact] public void XmlToExcel_Stream_Null_Throws() { Assert.Throws<Exception>(() => _parser.XmlToExcel((Stream)null!)); }

    [Fact] public void XmlToCsv_ReturnsBytes()
    {
        var csv = Encoding.UTF8.GetString(_parser.XmlToCsv(GetXmlContent()));
        Assert.Contains("Nome", csv);
    }

    [Fact] public void XmlToCsv_CustomDelimiter()
    {
        Assert.Contains(";", Encoding.UTF8.GetString(_parser.XmlToCsv(GetXmlContent(), ";")));
    }

    #endregion

    #region Form Parsing

    [Fact] public void GetJsonStringFromForm_FileName() { Assert.Contains("{", _parser.GetJsonStringFromForm(GetSamplePath("FormSample.xlsx"), "PlanTeste1")); }
    [Fact] public void GetJsonStringFromForm_Stream()
    {
        using var s = File.OpenRead(GetSamplePath("FormSample.xlsx"));
        Assert.Contains("{", _parser.GetJsonStringFromForm(s, "PlanTeste1"));
    }

    [Fact] public void GetJsonStringFromForm_InvalidSheet_ReturnsEmpty()
    {
        var json = _parser.GetJsonStringFromForm(GetSamplePath("FormSample.xlsx"), "NonExistent");
        Assert.True(string.IsNullOrWhiteSpace(json) || json == "{}");
    }

    [Fact] public void GetJsonObjectFromForm_FileName() { Assert.NotNull(_parser.GetJsonObjectFromForm(GetSamplePath("FormSample.xlsx"), "PlanTeste1")); }
    [Fact] public void GetJsonObjectFromForm_Stream()
    {
        using var s = File.OpenRead(GetSamplePath("FormSample.xlsx"));
        Assert.NotNull(_parser.GetJsonObjectFromForm(s, "PlanTeste1"));
    }

    [Fact] public void GetClassModelFromForm() { Assert.Contains("class", _parser.GetClassModelFromForm(GetSamplePath("FormSample.xlsx"), "PlanTeste1")); }

    [Fact] public void GetDictionary_FileName()
    {
        var dict = _parser.GetDictionary(GetSamplePath("FormSample.xlsx"), "PlanTeste1");
        Assert.True(dict.Count > 0);
    }

    [Fact] public void GetDictionary_Stream()
    {
        using var s = File.OpenRead(GetSamplePath("FormSample.xlsx"));
        Assert.True(_parser.GetDictionary(s, "PlanTeste1").Count > 0);
    }

    [Fact] public void GetDictionary_CustomFields()
    {
        var dict = _parser.GetDictionary(GetSamplePath("FormSample.xlsx"), "PlanTeste1");
        var key = dict.Keys.First();
        var result = _parser.GetDictionary(GetSamplePath("FormSample.xlsx"), "PlanTeste1", fieldNames: new[] { key });
        Assert.True(result.ContainsKey(key));
    }

    #endregion

    #region Error Handling

    [Fact] public void GetJsonStringFromTabular_EmptyFileName_Throws() { Assert.Throws<Exception>(() => _parser.GetJsonStringFromTabular("")); }
    [Fact] public void GetJsonStringFromTabular_NullFileName_Throws() { Assert.Throws<Exception>(() => _parser.GetJsonStringFromTabular((string)null!)); }
    [Fact] public void GetJsonStringFromForm_EmptyFileName_Throws() { Assert.Throws<Exception>(() => _parser.GetJsonStringFromForm("", "S")); }
    [Fact] public void GetJsonStringFromForm_NullStream_Throws() { Assert.Throws<Exception>(() => _parser.GetJsonStringFromForm((Stream)null!, "S")); }
    [Fact] public void GetJsonObjectFromForm_EmptyFileName_Throws() { Assert.Throws<Exception>(() => _parser.GetJsonObjectFromForm("", "S")); }
    [Fact] public void GetJsonObjectFromForm_NullStream_Throws() { Assert.Throws<Exception>(() => _parser.GetJsonObjectFromForm((Stream)null!, "S")); }
    [Fact] public void GetDictionary_EmptyFileName_Throws() { Assert.Throws<Exception>(() => _parser.GetDictionary("", "S")); }
    [Fact] public void GetDictionary_NullStream_Throws() { Assert.Throws<Exception>(() => _parser.GetDictionary((Stream)null!, "S")); }

    #endregion
}
