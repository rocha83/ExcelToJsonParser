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
