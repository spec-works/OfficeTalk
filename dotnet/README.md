# OfficeTalk .NET Library

A .NET 9 library for parsing, validating, and executing [OfficeTalk](https://spec-works.github.io/registry/parts/officetalk/) documents — deterministic modifications to Microsoft Office documents (Word, Excel, PowerPoint).

## Features

- **Parse** OfficeTalk documents into a typed AST
- **Validate** against syntactic rules and target documents
- **Execute** operations on Word (.docx) documents via OpenXML SDK
- **Address** elements using the OfficeTalk semantic path syntax

## Installation

```bash
dotnet add package SpecWorks.OfficeTalk
```

## Quick Start

### Parse an OfficeTalk Document

```csharp
using OfficeTalk.Parsing;

var source = @"
OFFICETALK/1.0
DOCTYPE word

AT body/heading[level=1]
SET ""Annual Report — FY2026""
FORMAT font-size=28pt, color=#1F3864

AT body/paragraph[text*=""teh company""]
REPLACE ""teh"" WITH ""the""
";

var lexer = new OfficeTalkLexer(source);
var tokens = lexer.Tokenize();
var parser = new OfficeTalkParser(tokens);
var document = parser.Parse();

Console.WriteLine($"Version: {document.Version}");
Console.WriteLine($"DocType: {document.DocType}");
Console.WriteLine($"Blocks: {document.OperationBlocks.Count}");
```

### Validate an OfficeTalk Document

```csharp
using OfficeTalk.Validation;

// Syntactic validation (no target document needed)
var syntacticValidator = new SyntacticValidator();
var syntacticResult = syntacticValidator.Validate(document);

if (!syntacticResult.IsValid)
{
    foreach (var error in syntacticResult.Errors)
        Console.Error.WriteLine(error);
}

// Semantic validation (against a target document)
var semanticValidator = new SemanticValidator();
var semanticResult = semanticValidator.Validate(document, "target.docx");
```

### Execute Operations on a Word Document

```csharp
using OfficeTalk.Execution;

var executor = new WordExecutor();
executor.Execute(document, "target.docx", "output.docx");
```

## Architecture

```
OfficeTalk/
├── Parsing/           Lexer and parser for OfficeTalk grammar
│   ├── Token.cs       Token types for the grammar
│   ├── OfficeTalkLexer.cs   Line-oriented tokenizer
│   └── OfficeTalkParser.cs  Parser producing typed AST
├── Ast/               Abstract syntax tree types
│   ├── OfficeTalkDocument.cs  Root AST node
│   ├── OperationBlock.cs      AT address + operations
│   ├── Address.cs             Address with segments
│   ├── AddressSegment.cs      Segment + predicates
│   ├── Predicate.cs           Predicate types
│   ├── Operations.cs          All operation types
│   └── DataTypes.cs           Color, Length, etc.
├── Addressing/        Resolve addresses against documents
│   ├── IAddressResolver.cs
│   └── WordAddressResolver.cs
├── Execution/         Execute operations on documents
│   ├── IOfficeTalkExecutor.cs
│   └── WordExecutor.cs
└── Validation/        Syntactic and semantic validation
    ├── SyntacticValidator.cs
    ├── SemanticValidator.cs
    └── ValidationResult.cs
```

### Processing Model

OfficeTalk uses a three-phase processing model:

1. **Resolution** — All addresses are resolved against the original document (snapshot semantics)
2. **Validation** — Semantic checks ensure operations are valid
3. **Execution** — Operations are applied sequentially

### Supported Operations

| Operation | Word | Excel | PowerPoint |
|-----------|------|-------|------------|
| SET | ✅ | — | — |
| REPLACE / REPLACE ALL | ✅ | — | — |
| INSERT BEFORE/AFTER | ◻️ | — | — |
| DELETE | ✅ | — | — |
| APPEND / PREPEND | ✅ | — | — |
| FORMAT | ◻️ | — | — |
| STYLE | ✅ | — | — |
| PROPERTY | ✅ | — | — |

✅ = Implemented, ◻️ = Stub (throws NotImplementedException), — = Not yet started

## Requirements

- .NET 9.0 or later
- DocumentFormat.OpenXml 3.1.0

## Testing

```bash
cd dotnet
dotnet test
```

## License

MIT License
