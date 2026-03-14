# OfficeTalk .NET Library

A .NET 9 library for parsing and validating [OfficeTalk](https://spec-works.github.io/registry/parts/officetalk/) documents — a grammar for deterministic modifications to Microsoft Office documents (Word, Excel, PowerPoint).

## Features

- **Parse** OfficeTalk documents into a typed AST
- **Validate** against syntactic rules (grammar-level checks, conflict detection)

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

var syntacticValidator = new SyntacticValidator();
var result = syntacticValidator.Validate(document);

if (!result.IsValid)
{
    foreach (var error in result.Errors)
        Console.Error.WriteLine(error);
}
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
└── Validation/        Syntactic validation
    ├── SyntacticValidator.cs
    └── ValidationResult.cs
```

## Requirements

- .NET 9.0 or later

## Testing

```bash
cd dotnet
dotnet test
```

## License

MIT License
