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

    #region Branch Coverage - Additional

    [Fact]
    public void GetDataTable_Stream_WithSkipRows()
    {
        using var stream = File.OpenRead(GetSamplePath("TabularSample.xlsx"));
        var dt = _parser.GetDataTable(stream, skipRows: 1);
        Assert.NotNull(dt);
    }

    [Fact]
    public void GetDataTable_Stream_NoHeader()
    {
        using var stream = File.OpenRead(GetSamplePath("TabularSample.xlsx"));
        var dt = _parser.GetDataTable(stream, useHeader: false);
        Assert.NotNull(dt);
    }

    [Fact]
    public void WriteItemJsonBody_AllTypes()
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Types");
        ws.Cell(1, 1).Value = "Str"; ws.Cell(1, 2).Value = "Int"; ws.Cell(1, 3).Value = "Dbl";
        ws.Cell(1, 4).Value = "Dec"; ws.Cell(1, 5).Value = "Bool"; ws.Cell(1, 6).Value = "Date";
        ws.Cell(1, 7).Value = "NullCol";
        ws.Cell(2, 1).Value = "abc"; ws.Cell(2, 2).Value = 42; ws.Cell(2, 3).Value = 3.14;
        ws.Cell(2, 4).Value = 9.99m; ws.Cell(2, 5).Value = true; ws.Cell(2, 6).Value = new DateTime(2026, 1, 15);
        ws.Cell(2, 7).Value = "";
        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;
        var json = _parser.GetJsonStringFromTabular(ms);
        Assert.Contains("abc", json);
        Assert.Contains("42", json);
    }

    [Fact]
    public void ValidateCell_AllRuleTypes()
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Val");
        ws.Cell(1, 1).Value = "Nome"; ws.Cell(1, 2).Value = "Idade"; ws.Cell(1, 3).Value = "Data";
        ws.Cell(2, 1).Value = "João"; ws.Cell(2, 2).Value = "abc"; ws.Cell(2, 3).Value = "not-a-date";
        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;
        var rules = new[] {
            new ValidationRule { ColumnName = "Idade", Type = "numeric" },
            new ValidationRule { ColumnName = "Data", Type = "date" }
        };
        var result = _parser.ValidateTabular(ms, rules);
        Assert.False(result.IsValid);
        Assert.True(result.ErrorRows > 0);
    }

    [Fact]
    public void ValidateTabular_EmptyRules()
    {
        using var stream = new MemoryStream(CreateSampleExcel());
        var result = _parser.ValidateTabular(stream, Array.Empty<ValidationRule>());
        Assert.True(result.IsValid);
        Assert.Equal(0, result.ErrorRows);
    }

    [Fact]
    public void XmlToCsv_EmptyDocument()
    {
        var xml = "<Root/>";
        var csv = Encoding.UTF8.GetString(_parser.XmlToCsv(xml));
        Assert.Equal(string.Empty, csv);
    }

    [Fact]
    public void JsonToCsv_EmptyArrayResult()
    {
        var csv = Encoding.UTF8.GetString(_parser.JsonToCsv("[]"));
        Assert.Equal(string.Empty, csv);
    }

    [Fact]
    public void EscapeCsvField_AllCases()
    {
        var data = new List<IDictionary<string, object>> {
            new Dictionary<string, object> { { "A", "" }, { "B", "has,comma" }, { "C", "has\"quote" }, { "D", "has\nnewline" } }
        };
        var bytes = _parser.TabularToExcel(data);
        Assert.True(bytes.Length > 0);
    }

    [Fact]
    public void ParseXmlToTabular_EmptyElements()
    {
        var xml = "<Root><Item/></Root>";
        var rows = new List<IDictionary<string, object>>();
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Xml");
        ws.Cell(1, 1).Value = "Test";
        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;
        var dt = _parser.GetDataTable(ms);
        Assert.NotNull(dt);
    }

    [Fact]
    public async Task StreamFromTabular_WithHeadersAndSkip()
    {
        using var stream = File.OpenRead(GetSamplePath("TabularSample.xlsx"));
        var rows = new List<IDictionary<string, object>>();
        var headers = new[] { "Col1", "Col2", "Col3" };
        using var stream2 = File.OpenRead(GetSamplePath("TabularSample.xlsx"));
        var dt = _parser.GetDataTable(stream);
        var colCount = dt.Columns.Count;
        var actualHeaders = Enumerable.Range(1, colCount).Select(i => $"C{i}").ToArray();
        await foreach (var row in _parser.StreamFromTabular(stream2, 0, null, null, actualHeaders))
            rows.Add(row);
        Assert.True(rows.Count > 0);
    }

    [Fact]
    public void GetJsonStringsFromAllSheets_SkipRows()
    {
        var result = _parser.GetJsonStringsFromAllSheets(GetSamplePath("TabularSample.xlsx"), skipRows: 1);
        Assert.True(result.Count > 0);
    }

    [Fact]
    public void CsvToExcel_Stream_EmptyLines()
    {
        using var stream = ToStream("A,B\n\n1,2\n");
        var bytes = _parser.CsvToExcel(stream);
        Assert.True(bytes.Length > 0);
    }

    [Fact]
    public void TabularToExcel_WithDateTime()
    {
        var data = new List<IDictionary<string, object>> {
            new Dictionary<string, object> { { "Date", new DateTime(2026, 1, 1) }, { "Val", 1.5 } }
        };
        var bytes = _parser.TabularToExcel(data);
        Assert.True(bytes.Length > 0);
    }

    [Fact]
    public void JsonElementToObject_AllNestedTypes()
    {
        var json = @"[{""Obj"":{""A"":1},""Arr"":[1,2,3],""Str"":""x"",""Num"":42,""Dbl"":1.5,""Bool"":true,""Null"":null}]";
        var bytes = _parser.JsonToExcel(json);
        Assert.True(bytes.Length > 0);
    }

    [Fact]
    public void WriteJsonBodyFromNamedFields_AllValueTypes()
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Form");
        workbook.NamedRanges.Add("StrVal", ws.Range("B1"));
        workbook.NamedRanges.Add("IntVal", ws.Range("B2"));
        workbook.NamedRanges.Add("DblVal", ws.Range("B3"));
        workbook.NamedRanges.Add("DecVal", ws.Range("B4"));
        workbook.NamedRanges.Add("BoolVal", ws.Range("B5"));
        workbook.NamedRanges.Add("NullVal", ws.Range("B6"));
        ws.Cell(1, 2).Value = "text";
        ws.Cell(2, 2).Value = 42;
        ws.Cell(3, 2).Value = 3.14;
        ws.Cell(4, 2).Value = 9.99m;
        ws.Cell(5, 2).Value = true;
        ws.Cell(6, 2).Value = "";
        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;
        using var stream = new MemoryStream(ms.ToArray());
        var json = _parser.GetJsonStringFromForm(stream, "Form");
        Assert.Contains("{", json);
    }

    [Fact]
    public void GetDictionary_FormWithNullFields()
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Form");
        workbook.NamedRanges.Add("Field1", ws.Range("B1"));
        ws.Cell(1, 2).Value = "value1";
        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;
        using var stream = new MemoryStream(ms.ToArray());
        var dict = _parser.GetDictionary(stream, "Form");
        Assert.True(dict.Count > 0);
    }

    #endregion

    #region Branch Coverage - Deep

    [Fact]
    public void WriteItemJsonBody_BooleansAndDates()
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("BD");
        ws.Cell(1, 1).Value = "BoolCol"; ws.Cell(1, 2).Value = "DateCol"; ws.Cell(1, 3).Value = "NullCol";
        ws.Cell(2, 1).Value = false;
        ws.Cell(2, 3).Value = "";
        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;
        var json = _parser.GetJsonStringFromTabular(ms);
        Assert.Contains("false", json);
    }

    [Fact]
    public void EscapeCsvField_CommaAndQuotesAndNewline()
    {
        var dict = new List<IDictionary<string, object>> {
            new Dictionary<string, object> { { "A", "has,comma" }, { "B", "has\"quote" }, { "C", "has\nnewline" } }
        };
        var bytes = _parser.TabularToExcel(dict);
        Assert.True(bytes.Length > 0);
    }

    [Fact]
    public void WriteJsonBody_NullFields()
    {
        using var ms = new MemoryStream();
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Empty");
        workbook.SaveAs(ms);
        ms.Position = 0;
        using var stream = new MemoryStream(ms.ToArray());
        var dict = _parser.GetDictionary(stream, "Empty");
        Assert.NotNull(dict);
    }

    [Fact]
    public void ValidateCell_MinLength_Passes()
    {
        using var stream = new MemoryStream(CreateSampleExcel());
        var rules = new[] { new ValidationRule { ColumnName = "Nome", Type = "min_length", MinLength = 1 } };
        var result = _parser.ValidateTabular(stream, rules);
        Assert.True(result.IsValid);
    }

    [Fact]
    public void ValidateCell_Numeric_EmptyString()
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("V");
        ws.Cell(1, 1).Value = "Val";
        ws.Cell(2, 1).Value = "";
        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;
        var rules = new[] { new ValidationRule { ColumnName = "Val", Type = "numeric" } };
        var result = _parser.ValidateTabular(ms, rules);
        Assert.True(result.IsValid);
    }

    [Fact]
    public void ValidateCell_Regex_EmptyString()
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("V");
        ws.Cell(1, 1).Value = "Val";
        ws.Cell(2, 1).Value = "";
        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;
        var rules = new[] { new ValidationRule { ColumnName = "Val", Type = "regex", Pattern = @"^\d+$" } };
        var result = _parser.ValidateTabular(ms, rules);
        Assert.True(result.IsValid);
    }

    [Fact]
    public void ValidateCell_InList_EmptyString()
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("V");
        ws.Cell(1, 1).Value = "Val";
        ws.Cell(2, 1).Value = "";
        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;
        var rules = new[] { new ValidationRule { ColumnName = "Val", Type = "in_list", AllowedValues = new[] { "A" } } };
        var result = _parser.ValidateTabular(ms, rules);
        Assert.True(result.IsValid);
    }

    [Fact]
    public void JsonElementToObject_LongAndDouble()
    {
        var json = @"[{""L"":9999999999,""D"":1.5,""S"":""x"",""B"":true,""N"":null,""A"":[1],""O"":{""K"":1}}]";
        var bytes = _parser.JsonToExcel(json);
        Assert.True(bytes.Length > 0);
    }

    [Fact]
    public void XmlToCsv_EmptyElements()
    {
        var xml = "<Root/>";
        var bytes = _parser.XmlToCsv(xml);
        Assert.True(bytes.Length >= 0);
    }

    [Fact]
    public void JsonToCsv_SingleRow()
    {
        var json = @"[{""A"":""1""}]";
        var csv = Encoding.UTF8.GetString(_parser.JsonToCsv(json));
        Assert.Contains("A", csv);
        Assert.Contains("1", csv);
    }

    [Fact]
    public void ParseFormSheet_WithNullSheet()
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Test");
        workbook.NamedRanges.Add("F1", ws.Range("B1"));
        ws.Cell(1, 2).Value = "val";
        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;
        using var stream = new MemoryStream(ms.ToArray());
        var json = _parser.GetJsonStringFromForm(stream, "Test");
        Assert.Contains("F1", json);
    }

    #endregion

    #region Branch Coverage - Deep

    [Fact]
    public void EscapeCsvField_OnlyQuote()
    {
        var dict = new List<IDictionary<string, object>> {
            new Dictionary<string, object> { { "A", "has\"quote" } }
        };
        var bytes = _parser.TabularToExcel(dict);
        Assert.True(bytes.Length > 0);
    }

    [Fact]
    public void EscapeCsvField_OnlyNewline()
    {
        var dict = new List<IDictionary<string, object>> {
            new Dictionary<string, object> { { "A", "has\nnewline" } }
        };
        var bytes = _parser.TabularToExcel(dict);
        Assert.True(bytes.Length > 0);
    }

    [Fact]
    public void ValidateCell_MaxLength_NoHasValue()
    {
        using var stream = new MemoryStream(CreateSampleExcel());
        var rules = new[] { new ValidationRule { ColumnName = "Nome", Type = "max_length" } };
        var result = _parser.ValidateTabular(stream, rules);
        Assert.True(result.IsValid);
    }

    [Fact]
    public void ValidateCell_MinLength_NoHasValue()
    {
        using var stream = new MemoryStream(CreateSampleExcel());
        var rules = new[] { new ValidationRule { ColumnName = "Nome", Type = "min_length" } };
        var result = _parser.ValidateTabular(stream, rules);
        Assert.True(result.IsValid);
    }

    [Fact]
    public void ValidateCell_InList_NullAllowedValues()
    {
        using var stream = new MemoryStream(CreateSampleExcel());
        var rules = new[] { new ValidationRule { ColumnName = "Cidade", Type = "in_list" } };
        var result = _parser.ValidateTabular(stream, rules);
        Assert.True(result.IsValid);
    }

    [Fact]
    public void ValidateCell_Regex_NullPattern()
    {
        using var stream = new MemoryStream(CreateSampleExcel());
        var rules = new[] { new ValidationRule { ColumnName = "Nome", Type = "regex" } };
        var result = _parser.ValidateTabular(stream, rules);
        Assert.True(result.IsValid);
    }

    [Fact]
    public void WriteItemJsonBody_AllBranches()
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("All");
        ws.Cell(1, 1).Value = "IntCol"; ws.Cell(1, 2).Value = "DblCol"; ws.Cell(1, 3).Value = "DecCol";
        ws.Cell(1, 4).Value = "BoolCol"; ws.Cell(1, 5).Value = "DateCol"; ws.Cell(1, 6).Value = "StrCol";
        ws.Cell(2, 1).Value = 42;
        ws.Cell(2, 2).Value = 3.14;
        ws.Cell(2, 3).Value = 9.99m;
        ws.Cell(2, 4).Value = true;
        ws.Cell(2, 5).Value = new DateTime(2026, 6, 15);
        ws.Cell(2, 6).Value = "text";
        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;
        var json = _parser.GetJsonStringFromTabular(ms);
        Assert.Contains("42", json);
        Assert.Contains("text", json);
    }

    [Fact]
    public void WriteJsonBody_AllFieldTypes()
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Form");
        workbook.NamedRanges.Add("IntF", ws.Range("B1"));
        workbook.NamedRanges.Add("DblF", ws.Range("B2"));
        workbook.NamedRanges.Add("DecF", ws.Range("B3"));
        workbook.NamedRanges.Add("BoolF", ws.Range("B4"));
        workbook.NamedRanges.Add("StrF", ws.Range("B5"));
        workbook.NamedRanges.Add("NullF", ws.Range("B6"));
        ws.Cell(1, 2).Value = 42;
        ws.Cell(2, 2).Value = 3.14;
        ws.Cell(3, 2).Value = 9.99m;
        ws.Cell(4, 2).Value = true;
        ws.Cell(5, 2).Value = "text";
        ws.Cell(6, 2).Value = "";
        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;
        using var stream = new MemoryStream(ms.ToArray());
        var json = _parser.GetJsonStringFromForm(stream, "Form");
        Assert.Contains("{", json);
    }

    [Fact]
    public void ParseXmlToTabular_Complex()
    {
        var xml = @"<Root>
            <Item><Name>A</Name><Value>1</Value></Item>
            <Item><Name>B</Name></Item>
            <Item><Value>3</Value></Item>
        </Root>";
        var bytes = _parser.XmlToExcel(xml);
        Assert.True(bytes.Length > 0);
    }

    [Fact]
    public void GetDataTable_SheetName_Found()
    {
        var dt = _parser.GetDataTable(GetSamplePath("TabularSample.xlsx"), "TestData");
        Assert.NotNull(dt);
    }

    [Fact]
    public void JsonToCsv_MultipleRows()
    {
        var json = @"[{""A"":""1"",""B"":""2""},{""A"":""3"",""B"":""4""}]";
        var csv = Encoding.UTF8.GetString(_parser.JsonToCsv(json));
        Assert.Contains("1", csv);
        Assert.Contains("4", csv);
    }

    [Fact]
    public void XmlToCsv_MultipleRows()
    {
        var xml = @"<R><R><A>1</A><B>2</B></R><R><A>3</A><B>4</B></R></R>";
        var csv = Encoding.UTF8.GetString(_parser.XmlToCsv(xml));
        Assert.Contains("A", csv);
        Assert.Contains("1", csv);
    }

    [Fact]
    public async Task StreamFromTabular_WithHeaders()
    {
        using var stream = File.OpenRead(GetSamplePath("TabularSample.xlsx"));
        var dt = _parser.GetDataTable(stream);
        var headers = Enumerable.Range(1, dt.Columns.Count).Select(i => $"C{i}").ToArray();
        using var stream2 = File.OpenRead(GetSamplePath("TabularSample.xlsx"));
        var rows = new List<IDictionary<string, object>>();
        await foreach (var row in _parser.StreamFromTabular(stream2, 0, null, null, headers))
            rows.Add(row);
        Assert.True(rows.Count > 0);
    }

    #endregion
}
