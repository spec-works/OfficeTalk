[!INCLUDE [](../dotnet/README.md)]

## Specification

The full OfficeTalk specification is maintained in [`officetalk-spec.md`](https://github.com/spec-works/OfficeTalk/blob/main/officetalk-spec.md).

Key sections:

- **§3 Document Structure** — Headers, operation blocks, inspect blocks
- **§4 Addressing** — Path syntax with segments and predicates
- **§5 Operations** — SET, DELETE, FORMAT, COMMENT, INSPECT, and more
- **§13 Formal Grammar** — Complete ABNF grammar
- **§14 Response Format** — JSONL response schema with CDDL definitions

## Test Cases

OfficeTalk includes a comprehensive test case suite in [`testcases/`](https://github.com/spec-works/OfficeTalk/tree/main/testcases):

- **Positive tests** — Valid OfficeTalk documents that must parse successfully
- **Negative tests** — Invalid documents that must produce parse errors

## API Reference

- [OfficeTalk API Documentation](api/OfficeTalk.html) - Parser, AST, and validator API reference
