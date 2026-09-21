# Rochas.ExcelToJsonParser

[English](#english) | [Português](#português) | [Español](#español) | [Français](#français) | [Deutsch](#deutsch)

---

## English

# README - ExcelToJsonParser

Utility component for converting Excel files to JSON, DataTable, dynamic Objects and C# Models (automatically generated via NJsonSchema).

It supports two main reading modes:

1. **Tabular Sheet (table-format spreadsheets)**
2. **Form Sheet (form-structured spreadsheets)**

## Main Features

## Tabular Mode Reading
Table-format spreadsheets (rows x columns)

JSON as string

```csharp
var parser = new ExcelToJsonParser();
string json = parser.GetJsonStringFromTabular("arquivo.xlsx");
```

---

JSON as objects (IEnumerable<object>)

```csharp
var parser = new ExcelToJsonParser();
var objList = parser.GetJsonObjectFromTabular("arquivo.xlsx");
foreach(var obj in objList)
{
	...
}
```

---

DataTable (with or without header)

```csharp
DataTable data = parser.GetDataTable("arquivo.xlsx", skipRows: 1, useHeader: true);
```

---

C# classes from column names

```csharp
string classFile = parser.GetClassModelFromTabular("arquivo.xlsx");
```

---

## Form Mode
Form-structured spreadsheets (e.g.: "Field: Value").

JSON as string

```csharp
string json = parser.GetJsonStringFromForm("arquivo.xlsx", "FichaCliente");
```

---

JSON as object

```csharp
var obj = parser.GetJsonObjectFromForm("arquivo.xlsx", "FichaCliente");
```

---

Dictionary<string, object>

```csharp
var dict = parser.GetDictionary("arquivo.xlsx", "FichaCliente");
```

---

C# class

```csharp
string classModel = parser.GetClassModelFromForm("arquivo.xlsx", "FichaCliente");
```

---

## Important Parameters

**skipRows**:
Skips leading rows.

**replaceFrom / replaceTo**:
Replaces parts of column names.

**headerColumns**:
Manually provides the Excel header.

**onlySampleRow**:
When true, reads only 1 row.
Used internally for C# model generation.

---

## Usage Examples
Read a tabular spreadsheet skipping 2 rows and normalizing headers

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

## Returned Structure

Typical Tabular mode example:

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

Form Mode:

```json
{
  "Nome": "Carlos",
  "CPF": "111.222.333-44",
  "Telefone": "(11) 99999-0000"
}
```

---

## Streaming (IAsyncEnumerable)

Row-by-row async processing for large data volumes.

```csharp
using var parser = new ExcelToJsonParser();
using var stream = File.OpenRead("grande.xlsx");

await foreach (var row in parser.StreamFromTabular(stream))
{
    Console.WriteLine(row["Nome"]);
}
```

---

## Validation

Per-column validation rules before processing.

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

Available types: `required`, `max_length`, `min_length`, `numeric`, `date`, `in_list`, `regex`.

---

## Transformation

Apply transformations to the read values.

```csharp
var transforms = new[]
{
    new TransformConfig { Type = "trim" },
    new TransformConfig { Type = "upper_case" },
    new TransformConfig { Type = "replace", Params = new() { { "from", " " }, { "to", "_" } } }
};

var clean = parser.ApplyTransforms("  joao silva  ", transforms); // "JOAO_SILVA"
```

Available types: `trim`, `upper_case`, `lower_case`, `title_case`, `replace`, `to_decimal`, `to_int`, `to_date`, `to_boolean`, `default`, `split`, `map_values`.

---

## Multi-sheet

Read all sheets at once.

```csharp
var allSheets = parser.GetJsonStringsFromAllSheets("multi.xlsx");
foreach (var kvp in allSheets)
    Console.WriteLine($"{kvp.Key}: {kvp.Value}");
```

---

## Export (Data → Excel/CSV)

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

## Reverse Flow (Data ← JSON/CSV/XML)

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

## IDisposable

`ExcelToJsonParser` implements `IDisposable`. Use `using` to ensure proper resource disposal.

```csharp
using var parser = new ExcelToJsonParser();
// ... usage
```

---

## Dependencies

- **ExcelDataReader** 3.9.0 — .xls/.xlsx/.xlsb reading
- **ClosedXML** 0.102.3 — .xlsx writing
- **System.Text.Json** 8.0.5 — JSON serialization
- **NJsonSchema** 10.1.16 — C# class generation

---

## Tests

```bash
dotnet test Rochas.ExcelToJsonParser.Tests/
```

140 tests covering: Tabular, Form, Streaming, Multi-sheet, Validation, Transform, Export, Reverse Flow, Error Handling.

---

## Support

- .NET Standard 2.1 / .NET 6+ / .NET 8+ / .NET 9
- NuGet: `dotnet add package Rochas.ExcelToJsonParser`

## Português

# README - ExcelToJsonParser

Componente utilitário para conversão de arquivos Excel em JSON, DataTable, Objetos dinâmicos e Modelos C# (gerados automaticamente via NJsonSchema).

Ele suporta dois modos principais de leitura:

1. **Tabular Sheet (planilhas em formato de tabela)**
2. **Form Sheet (planilhas estruturadas como formulários)**

## Funcionalidades Principais

## Leitura em Modo Tabular
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

## Form Mode
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

## Parâmetros Importantes

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

## Exemplos de Uso
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

## Estrutura Retornada

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

## Streaming (IAsyncEnumerable)

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

## Validação

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

## Transformação

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

## Multi-sheet

Ler todas as planilhas de uma vez.

```csharp
var allSheets = parser.GetJsonStringsFromAllSheets("multi.xlsx");
foreach (var kvp in allSheets)
    Console.WriteLine($"{kvp.Key}: {kvp.Value}");
```

---

## Exportação (Dados → Excel/CSV)

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

## Fluxo Inverso (Dados ← JSON/CSV/XML)

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

## IDisposable

O `ExcelToJsonParser` implementa `IDisposable`. Use `using` para garantir liberação adequada de recursos.

```csharp
using var parser = new ExcelToJsonParser();
// ... uso
```

---

## Dependências

- **ExcelDataReader** 3.9.0 — Leitura .xls/.xlsx/.xlsb
- **ClosedXML** 0.102.3 — Escrita .xlsx
- **System.Text.Json** 8.0.5 — Serialização JSON
- **NJsonSchema** 10.1.16 — Geração de classes C#

---

## Testes

```bash
dotnet test Rochas.ExcelToJsonParser.Tests/
```

140 testes cobrindo: Tabular, Form, Streaming, Multi-sheet, Validation, Transform, Export, Reverse Flow, Error Handling.

---

## Suporte

- .NET Standard 2.1 / .NET 6+ / .NET 8+ / .NET 9
- NuGet: `dotnet add package Rochas.ExcelToJsonParser`

## Español

# README - ExcelToJsonParser

Componente utilitario para convertir archivos Excel a JSON, DataTable, Objetos dinámicos y Modelos C# (generados automáticamente vía NJsonSchema).

Soporta dos modos principales de lectura:

1. **Tabular Sheet (hojas en formato de tabla)**
2. **Form Sheet (hojas estructuradas como formularios)**

## Funcionalidades Principales

## Lectura en Modo Tabular
Hojas en formato tabla (filas x columnas)

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

DataTable (con o sin encabezado)

```csharp
DataTable data = parser.GetDataTable("arquivo.xlsx", skipRows: 1, useHeader: true);
```

---

Clases C# a partir de los nombres de las columnas

```csharp
string classFile = parser.GetClassModelFromTabular("arquivo.xlsx");
```

---

## Form Mode
Hojas estructuradas como formulario (ej.: "Campo: Valor").

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

Clase C#

```csharp
string classModel = parser.GetClassModelFromForm("arquivo.xlsx", "FichaCliente");
```

---

## Parámetros Importantes

**skipRows**:
Omite las filas iniciales.

**replaceFrom / replaceTo**:
Permite reemplazar partes de los nombres de las columnas.

**headerColumns**:
Permite informar manualmente el encabezado de Excel.

**onlySampleRow**:
Cuando es true, lee solo 1 fila.
Usado internamente para la generación de modelos C#.

---

## Ejemplos de Uso
Leer hoja tabular omitiendo 2 filas y normalizando encabezados

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

## Estructura Devuelta

Ejemplo típico del modo Tabular:

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

## Streaming (IAsyncEnumerable)

Procesamiento asíncrono fila por fila para grandes volúmenes de datos.

```csharp
using var parser = new ExcelToJsonParser();
using var stream = File.OpenRead("grande.xlsx");

await foreach (var row in parser.StreamFromTabular(stream))
{
    Console.WriteLine(row["Nome"]);
}
```

---

## Validación

Reglas de validación por columna antes de procesar.

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

Tipos disponibles: `required`, `max_length`, `min_length`, `numeric`, `date`, `in_list`, `regex`.

---

## Transformación

Aplicar transformaciones a los valores leídos.

```csharp
var transforms = new[]
{
    new TransformConfig { Type = "trim" },
    new TransformConfig { Type = "upper_case" },
    new TransformConfig { Type = "replace", Params = new() { { "from", " " }, { "to", "_" } } }
};

var clean = parser.ApplyTransforms("  joao silva  ", transforms); // "JOAO_SILVA"
```

Tipos disponibles: `trim`, `upper_case`, `lower_case`, `title_case`, `replace`, `to_decimal`, `to_int`, `to_date`, `to_boolean`, `default`, `split`, `map_values`.

---

## Multi-sheet

Leer todas las hojas de una vez.

```csharp
var allSheets = parser.GetJsonStringsFromAllSheets("multi.xlsx");
foreach (var kvp in allSheets)
    Console.WriteLine($"{kvp.Key}: {kvp.Value}");
```

---

## Exportación (Datos → Excel/CSV)

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

## Flujo Inverso (Datos ← JSON/CSV/XML)

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

## IDisposable

`ExcelToJsonParser` implementa `IDisposable`. Use `using` para garantizar la liberación adecuada de recursos.

```csharp
using var parser = new ExcelToJsonParser();
// ... uso
```

---

## Dependencias

- **ExcelDataReader** 3.9.0 — Lectura .xls/.xlsx/.xlsb
- **ClosedXML** 0.102.3 — Escritura .xlsx
- **System.Text.Json** 8.0.5 — Serialización JSON
- **NJsonSchema** 10.1.16 — Generación de clases C#

---

## Pruebas

```bash
dotnet test Rochas.ExcelToJsonParser.Tests/
```

140 pruebas cubriendo: Tabular, Form, Streaming, Multi-sheet, Validation, Transform, Export, Reverse Flow, Error Handling.

---

## Soporte

- .NET Standard 2.1 / .NET 6+ / .NET 8+ / .NET 9
- NuGet: `dotnet add package Rochas.ExcelToJsonParser`

## Français

# README - ExcelToJsonParser

Composant utilitaire pour convertir des fichiers Excel en JSON, DataTable, Objets dynamiques et Modèles C# (générés automatiquement via NJsonSchema).

Il prend en charge deux modes de lecture principaux :

1. **Tabular Sheet (feuilles au format tableau)**
2. **Form Sheet (feuilles structurées comme des formulaires)**

## Fonctionnalités Principales

## Lecture en Mode Tabular
Feuilles au format tableau (lignes x colonnes)

JSON en tant que string

```csharp
var parser = new ExcelToJsonParser();
string json = parser.GetJsonStringFromTabular("arquivo.xlsx");
```

---

JSON en tant qu'objets (IEnumerable<object>)

```csharp
var parser = new ExcelToJsonParser();
var objList = parser.GetJsonObjectFromTabular("arquivo.xlsx");
foreach(var obj in objList)
{
	...
}
```

---

DataTable (avec ou sans en-tête)

```csharp
DataTable data = parser.GetDataTable("arquivo.xlsx", skipRows: 1, useHeader: true);
```

---

Classes C# à partir des noms de colonnes

```csharp
string classFile = parser.GetClassModelFromTabular("arquivo.xlsx");
```

---

## Form Mode
Feuilles structurées comme un formulaire (ex. : "Champ: Valeur").

JSON en tant que string

```csharp
string json = parser.GetJsonStringFromForm("arquivo.xlsx", "FichaCliente");
```

---

JSON en tant qu'objet

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

## Paramètres Importants

**skipRows** :
Ignore les premières lignes.

**replaceFrom / replaceTo** :
Permet de remplacer des parties des noms de colonnes.

**headerColumns** :
Permet de fournir manuellement l'en-tête Excel.

**onlySampleRow** :
Quand true, ne lit qu'une seule ligne.
Utilisé en interne pour la génération de modèles C#.

---

## Exemples d'Utilisation
Lire une feuille tabulaire en ignorant 2 lignes et en normalisant les en-têtes

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

## Structure Retournée

Exemple typique du mode Tabular :

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

Mode Form :

```json
{
  "Nome": "Carlos",
  "CPF": "111.222.333-44",
  "Telefone": "(11) 99999-0000"
}
```

---

## Streaming (IAsyncEnumerable)

Traitement asynchrone ligne par ligne pour de gros volumes de données.

```csharp
using var parser = new ExcelToJsonParser();
using var stream = File.OpenRead("grande.xlsx");

await foreach (var row in parser.StreamFromTabular(stream))
{
    Console.WriteLine(row["Nome"]);
}
```

---

## Validation

Règles de validation par colonne avant traitement.

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

Types disponibles : `required`, `max_length`, `min_length`, `numeric`, `date`, `in_list`, `regex`.

---

## Transformation

Appliquer des transformations aux valeurs lues.

```csharp
var transforms = new[]
{
    new TransformConfig { Type = "trim" },
    new TransformConfig { Type = "upper_case" },
    new TransformConfig { Type = "replace", Params = new() { { "from", " " }, { "to", "_" } } }
};

var clean = parser.ApplyTransforms("  joao silva  ", transforms); // "JOAO_SILVA"
```

Types disponibles : `trim`, `upper_case`, `lower_case`, `title_case`, `replace`, `to_decimal`, `to_int`, `to_date`, `to_boolean`, `default`, `split`, `map_values`.

---

## Multi-sheet

Lire toutes les feuilles à la fois.

```csharp
var allSheets = parser.GetJsonStringsFromAllSheets("multi.xlsx");
foreach (var kvp in allSheets)
    Console.WriteLine($"{kvp.Key}: {kvp.Value}");
```

---

## Exportation (Données → Excel/CSV)

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

## Flux Inverse (Données ← JSON/CSV/XML)

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

## IDisposable

`ExcelToJsonParser` implémente `IDisposable`. Utilisez `using` pour garantir une libération adéquate des ressources.

```csharp
using var parser = new ExcelToJsonParser();
// ... usage
```

---

## Dépendances

- **ExcelDataReader** 3.9.0 — Lecture .xls/.xlsx/.xlsb
- **ClosedXML** 0.102.3 — Écriture .xlsx
- **System.Text.Json** 8.0.5 — Sérialisation JSON
- **NJsonSchema** 10.1.16 — Génération de classes C#

---

## Tests

```bash
dotnet test Rochas.ExcelToJsonParser.Tests/
```

140 tests couvrant : Tabular, Form, Streaming, Multi-sheet, Validation, Transform, Export, Reverse Flow, Error Handling.

---

## Support

- .NET Standard 2.1 / .NET 6+ / .NET 8+ / .NET 9
- NuGet : `dotnet add package Rochas.ExcelToJsonParser`

## Deutsch

# README - ExcelToJsonParser

Dienstprogrammkomponente zum Konvertieren von Excel-Dateien in JSON, DataTable, dynamische Objekte und C#-Modelle (automatisch generiert via NJsonSchema).

Es werden zwei Hauptlesemodi unterstützt:

1. **Tabular Sheet (Tabellenformat)**
2. **Form Sheet (formularstrukturierte Blätter)**

## Hauptfunktionen

## Lesen im Tabular-Modus
Blätter im Tabellenformat (Zeilen x Spalten)

JSON als String

```csharp
var parser = new ExcelToJsonParser();
string json = parser.GetJsonStringFromTabular("arquivo.xlsx");
```

---

JSON als Objekte (IEnumerable<object>)

```csharp
var parser = new ExcelToJsonParser();
var objList = parser.GetJsonObjectFromTabular("arquivo.xlsx");
foreach(var obj in objList)
{
	...
}
```

---

DataTable (mit oder ohne Kopfzeile)

```csharp
DataTable data = parser.GetDataTable("arquivo.xlsx", skipRows: 1, useHeader: true);
```

---

C#-Klassen aus Spaltennamen

```csharp
string classFile = parser.GetClassModelFromTabular("arquivo.xlsx");
```

---

## Form Mode
Blätter als Formular strukturiert (z. B.: "Feld: Wert").

JSON als String

```csharp
string json = parser.GetJsonStringFromForm("arquivo.xlsx", "FichaCliente");
```

---

JSON als Objekt

```csharp
var obj = parser.GetJsonObjectFromForm("arquivo.xlsx", "FichaCliente");
```

---

Dictionary<string, object>

```csharp
var dict = parser.GetDictionary("arquivo.xlsx", "FichaCliente");
```

---

C#-Klasse

```csharp
string classModel = parser.GetClassModelFromForm("arquivo.xlsx", "FichaCliente");
```

---

## Wichtige Parameter

**skipRows**:
Überspringt die ersten Zeilen.

**replaceFrom / replaceTo**:
Ersetzt Teile der Spaltennamen.

**headerColumns**:
Excel-Kopfzeile manuell angeben.

**onlySampleRow**:
Wenn true, wird nur 1 Zeile gelesen.
Wird intern für die C#-Modellgenerierung verwendet.

---

## Verwendungsbeispiele
Tabellenblatt lesen, dabei 2 Zeilen überspringen und Kopfzeilen normalisieren

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

## Rückgabestruktur

Typisches Beispiel für den Tabular-Modus:

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

Form-Modus:

```json
{
  "Nome": "Carlos",
  "CPF": "111.222.333-44",
  "Telefone": "(11) 99999-0000"
}
```

---

## Streaming (IAsyncEnumerable)

Asynchrone zeilenweise Verarbeitung für große Datenmengen.

```csharp
using var parser = new ExcelToJsonParser();
using var stream = File.OpenRead("grande.xlsx");

await foreach (var row in parser.StreamFromTabular(stream))
{
    Console.WriteLine(row["Nome"]);
}
```

---

## Validierung

Spaltenweise Validierungsregeln vor der Verarbeitung.

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

Verfügbare Typen: `required`, `max_length`, `min_length`, `numeric`, `date`, `in_list`, `regex`.

---

## Transformation

Transformationen auf die gelesenen Werte anwenden.

```csharp
var transforms = new[]
{
    new TransformConfig { Type = "trim" },
    new TransformConfig { Type = "upper_case" },
    new TransformConfig { Type = "replace", Params = new() { { "from", " " }, { "to", "_" } } }
};

var clean = parser.ApplyTransforms("  joao silva  ", transforms); // "JOAO_SILVA"
```

Verfügbare Typen: `trim`, `upper_case`, `lower_case`, `title_case`, `replace`, `to_decimal`, `to_int`, `to_date`, `to_boolean`, `default`, `split`, `map_values`.

---

## Multi-sheet

Alle Blätter auf einmal lesen.

```csharp
var allSheets = parser.GetJsonStringsFromAllSheets("multi.xlsx");
foreach (var kvp in allSheets)
    Console.WriteLine($"{kvp.Key}: {kvp.Value}");
```

---

## Export (Daten → Excel/CSV)

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

## Umgekehrter Fluss (Daten ← JSON/CSV/XML)

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

## IDisposable

`ExcelToJsonParser` implementiert `IDisposable`. `using` verwenden, um eine ordnungsgemäße Ressourcenfreigabe sicherzustellen.

```csharp
using var parser = new ExcelToJsonParser();
// ... Verwendung
```

---

## Abhängigkeiten

- **ExcelDataReader** 3.9.0 — Lesen .xls/.xlsx/.xlsb
- **ClosedXML** 0.102.3 — Schreiben .xlsx
- **System.Text.Json** 8.0.5 — JSON-Serialisierung
- **NJsonSchema** 10.1.16 — C#-Klassengenerierung

---

## Tests

```bash
dotnet test Rochas.ExcelToJsonParser.Tests/
```

140 Tests für: Tabular, Form, Streaming, Multi-sheet, Validation, Transform, Export, Reverse Flow, Error Handling.

---

## Support

- .NET Standard 2.1 / .NET 6+ / .NET 8+ / .NET 9
- NuGet: `dotnet add package Rochas.ExcelToJsonParser`
