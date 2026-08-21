# README - ExcelToJsonParser

Componente utilitário para conversão de arquivos Excel em JSON, DataTable, Objetos dinâmicos e Modelos C# (gerados automaticamente via NJsonSchema).

Ele suporta dois modos principais de leitura:

1. **Tabular Sheet (planilhas em formato de tabela)**  
2. **Form Sheet (planilhas estruturadas como formulários)**  

## 🚀 Funcionalidades Principais

## ✔ Leitura em Modo Tabular
Planilhas no formato tabela (linhas x colunas)

JSON como string

```csharp
var parser = new ExcelToJsonParser();
string json = parser.GetJsonStringFromTabular("arquivo.xlsx");
```

---

JSON como objetos (IEnumerable<object>)

```csharp
var parser = new ExcelToJsonParser();
var objList = parser.GetJsonObjectFromTabular("arquivo.xlsx");
foreach(var obj in objList)
{
	...
}
```

---

DataTable (com ou sem cabeçalho)

```csharp
DataTable data = parser.GetDataTable("arquivo.xlsx", skipRows: 1, useHeader: true);
```

---

Classes C# a partir dos nomes das colunas

```csharp
string classFile = parser.GetClassModelFromTabular("arquivo.xlsx");
```

---

## ✔ Form Mode
Planilhas estruturadas como formulário (ex.: "Campo: Valor").

JSON como string

```csharp
string json = parser.GetJsonStringFromForm("arquivo.xlsx", "FichaCliente");
```

---

JSON como objeto

```csharp
var obj = parser.GetJsonObjectFromForm("arquivo.xlsx", "FichaCliente");
```

---

Dictionary<string, object>

```csharp
var dict = parser.GetDictionary("arquivo.xlsx", "FichaCliente");
```

---

Classe C#

```csharp
string classModel = parser.GetClassModelFromForm("arquivo.xlsx", "FichaCliente");
```

---

## 🎯 Parâmetros Importantes

**skipRows**:
Ignora linhas iniciais.

**replaceFrom / replaceTo**:
Permite substituir partes do nome das colunas.

**headerColumns**:
Permite informar manualmente o cabeçalho do Excel.

**onlySampleRow**:
Quando true, lê apenas 1 linha.
Usado internamente para geração de modelos C#.

---

## 🔧 Exemplos de Uso
Ler planilha tabular ignorando 2 linhas e normalizando cabeçalhos

```csharp
var parser = new ExcelToJsonParser();

string json = parser.GetJsonStringFromTabular(
    "produtos.xlsx",
    skipRows: 2,
    replaceFrom: new[] { " ", "-" },
    replaceTo:   new[] { "_", "" }
);

Console.WriteLine(json);
```

---

## 🧱 Estrutura Retornada

Exemplo típico do modo Tabular:

```json
[
  {
    "Nome": "Ana",
    "Idade": 30,
    "Ativo": true
  },
  {
    "Nome": "João",
    "Idade": 22,
    "Ativo": false
  }
]
``` 

---

Modo Form:

```json
{
  "Nome": "Carlos",
  "CPF": "111.222.333-44",
  "Telefone": "(11) 99999-0000"
}
``` 

---

## ✔ Streaming (IAsyncEnumerable)

Processamento assíncrono linha a linha para grandes volumes de dados.

```csharp
using var parser = new ExcelToJsonParser();
using var stream = File.OpenRead("grande.xlsx");

await foreach (var row in parser.StreamFromTabular(stream))
{
    Console.WriteLine(row["Nome"]);
}
```

---

## ✔ Validação

Regras de validação por coluna antes de processar.

```csharp
var rules = new[]
{
    new ValidationRule { ColumnName = "Email", Type = "required", Message = "Email obrigatório" },
    new ValidationRule { ColumnName = "Idade", Type = "numeric" },
    new ValidationRule { ColumnName = "UF", Type = "in_list", AllowedValues = new[] { "SP", "RJ" } },
    new ValidationRule { ColumnName = "CPF", Type = "regex", Pattern = @"^\d{3}\.\d{3}\.\d{3}-\d{2}$" }
};

var result = parser.ValidateTabular(stream, rules);
if (!result.IsValid)
    result.Errors.ForEach(e => Console.WriteLine(e.Message));
```

Tipos disponíveis: `required`, `max_length`, `min_length`, `numeric`, `date`, `in_list`, `regex`.

---

## ✔ Transformação

Aplicar transformações nos valores lidos.

```csharp
var transforms = new[]
{
    new TransformConfig { Type = "trim" },
    new TransformConfig { Type = "upper_case" },
    new TransformConfig { Type = "replace", Params = new() { { "from", " " }, { "to", "_" } } }
};

var clean = parser.ApplyTransforms("  joao silva  ", transforms); // "JOAO_SILVA"
```

Tipos disponíveis: `trim`, `upper_case`, `lower_case`, `title_case`, `replace`, `to_decimal`, `to_int`, `to_date`, `to_boolean`, `default`, `split`, `map_values`.

---

## ✔ Multi-sheet

Ler todas as planilhas de uma vez.

```csharp
var allSheets = parser.GetJsonStringsFromAllSheets("multi.xlsx");
foreach (var kvp in allSheets)
    Console.WriteLine($"{kvp.Key}: {kvp.Value}");
```

---

## ✔ Exportação (Dados → Excel/CSV)

### TabularToExcel
```csharp
var data = new List<IDictionary<string, object>>
{
    new Dictionary<string, object> { { "Nome", "João" }, { "Idade", 30 } }
};
byte[] excelBytes = parser.TabularToExcel(data, "Pessoas");
File.WriteAllBytes("saida.xlsx", excelBytes);
```

### CsvToExcel
```csharp
using var csvStream = File.OpenRead("dados.csv");
byte[] excelBytes = parser.CsvToExcel(csvStream, delimiter: ",");
File.WriteAllBytes("saida.xlsx", excelBytes);
```

---

## ✔ Fluxo Inverso (Dados ← JSON/CSV/XML)

### JSON → Excel
```csharp
var json = @"[{""Nome"":""João"",""Idade"":30},{""Nome"":""Maria"",""Idade"":25}]";
byte[] excelBytes = parser.JsonToExcel(json, "Pessoas");
File.WriteAllBytes("saida.xlsx", excelBytes);
```

### JSON → CSV
```csharp
byte[] csvBytes = parser.JsonToCsv(json, delimiter: ";");
File.WriteAllBytes("saida.csv", csvBytes);
```

### JSON → CSV (Stream)
```csharp
using Stream csvStream = parser.JsonToCsvStream(json);
```

### XML → Excel
```csharp
var xml = @"<Pessoas><Pessoa><Nome>João</Nome><Idade>30</Idade></Pessoa></Pessoas>";
byte[] excelBytes = parser.XmlToExcel(xml, "Pessoas");
File.WriteAllBytes("saida.xlsx", excelBytes);
```

### XML → CSV
```csharp
byte[] csvBytes = parser.XmlToCsv(xml);
```

---

## ♻️ IDisposable

O `ExcelToJsonParser` implementa `IDisposable`. Use `using` para garantir liberação adequada de recursos.

```csharp
using var parser = new ExcelToJsonParser();
// ... uso
```

---

## 📦 Dependências

- **ExcelDataReader** 3.9.0 — Leitura .xls/.xlsx/.xlsb
- **ClosedXML** 0.102.3 — Escrita .xlsx
- **System.Text.Json** 8.0.5 — Serialização JSON
- **NJsonSchema** 10.1.16 — Geração de classes C#

---

## 🧪 Testes

```bash
dotnet test Rochas.ExcelToJsonParser.Tests/
```

102 testes cobrindo: Tabular, Form, Streaming, Multi-sheet, Validation, Transform, Export, Reverse Flow, Error Handling.

---

## Suporte

- .NET Standard 2.1 / .NET 6+ / .NET 8+ / .NET 9
- NuGet: `dotnet add package Rochas.ExcelToJsonParser`
