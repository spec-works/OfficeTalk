# OfficeTalk Inspect Operations and JSONL Response Format

**Proposal for OfficeTalk/1.0 Specification Update**

## Motivation

OfficeTalk defines a powerful write path for deterministic document modification.
However, practical use — especially by LLM agents — requires a corresponding
**read path**. An agent cannot reliably construct AT addresses and operations
without first understanding what the target document contains.

The current CLI `inspect` command fills this gap informally, but the response
format is unspecified, implementation-specific, and human-oriented. This proposal
formalizes:

1. **INSPECT** — a read operation within the OfficeTalk document grammar.
2. **JSONL Response Format** — a structured, streamable response format for
   both INSPECT results and operation execution feedback.

## Design Principles

- **Same grammar, same addressing** — INSPECT reuses the existing OfficeTalk
  address syntax. No new addressing concepts.
- **Read documents are separate from write documents** — An `.otk` document
  contains either INSPECT blocks or AT/operation blocks, never both.
- **Layered detail** — Callers control how much information is returned:
  addressing (structure), content (values), and properties (formatting).
- **Streamable** — JSONL (one JSON object per line) allows responses to be
  streamed as each INSPECT or operation is processed.
- **Atomic execution is preserved** — Write operations remain atomic. The
  response format is an optional reporting layer, not a change to execution
  semantics.

---

## §5.6. Inspect Operations

### 5.6.1. INSPECT

Resolves an address against the target document and returns a structured
description of the matched elements.

```
INSPECT <address>
  [DEPTH <integer>]
  [INCLUDE <layer> *("," <layer>)]
  [CONTEXT <integer>]
```

**Parameters:**

| Keyword   | Type    | Default | Description |
|-----------|---------|---------|-------------|
| `DEPTH`   | integer | 0       | Levels of child elements to include. 0 = matched element only. |
| `INCLUDE` | layers  | (none)  | Comma-separated list: `content`, `properties`, or both. |
| `CONTEXT` | integer | 0       | Number of sibling elements before and after to include. |

When no INCLUDE is specified, only addressing information is returned (element
type, position, identity). This is the most lightweight mode — suitable for
confirming an address resolves correctly or mapping document structure.

### 5.6.2. Examples

```
OFFICETALK/1.0
DOCTYPE excel

# Discover what sheets exist (addressing only)
INSPECT sheet[1]

# See all rows in the Q1 Budget sheet with their values
INSPECT sheet["Q1 Budget"]
  DEPTH 1
  INCLUDE content

# Get formatting details for a specific cell
INSPECT sheet["Q1 Budget"]/D2
  INCLUDE content, properties

# See a cell with its neighbors
INSPECT sheet["Q1 Budget"]/D2
  INCLUDE content
  CONTEXT 2
```

```
OFFICETALK/1.0
DOCTYPE word

# Get the document outline — all headings
INSPECT body/heading
  INCLUDE content

# Inspect a specific heading with its surrounding paragraphs
INSPECT body/heading[text="Conclusion"]
  INCLUDE content
  CONTEXT 3

# See a table's full structure with formatting
INSPECT body/table[1]
  DEPTH 2
  INCLUDE content, properties
```

```
OFFICETALK/1.0
DOCTYPE powerpoint

# Outline all slides (titles only)
INSPECT slide
  DEPTH 1
  INCLUDE content

# Full detail on slide 3 including comments
INSPECT slide[3]
  DEPTH 1
  INCLUDE content, properties
```

### 5.6.3. INSPECT Semantics

An OfficeTalk document containing INSPECT operations is a **read document**.
It MUST NOT contain AT blocks, PROPERTY statements, or any write operations.
Implementations MUST reject documents that mix INSPECT and write operations.

Multiple INSPECT operations in a single document are processed sequentially.
Each INSPECT produces one result object in the response.

INSPECT operations do not modify the target document.

### 5.6.4. Detail Layers

The INCLUDE keyword controls which information layers are returned for each
matched element. Layers are additive.

**Addressing (always included):**

Structural identity of the element — type, position within its container,
and type-specific identifiers (sheet name, cell reference, placeholder type,
heading level, style name).

**Content (INCLUDE content):**

The textual content of the element. For cells, the display value. For
paragraphs, the full text. For shapes, the text content. For rows, the
cell values. Content is not included by default to keep responses lightweight
when only structure is needed.

**Properties (INCLUDE properties):**

Formatting and metadata properties of the element. Font name, size, bold,
italic, color, fill, borders, number format, alignment. Properties are
expensive to compute and verbose in output — only requested when the caller
needs to reason about or modify formatting.

---

## §14. Response Format

### 14.1. Overview

OfficeTalk defines a JSONL (JSON Lines) response format for communicating
results back to callers. Each line in the response is a self-contained JSON
object terminated by a newline (`\n`).

The response format applies to:

- **INSPECT operations** — each INSPECT produces one response object
  describing the matched elements.
- **Write operations** — each AT block optionally produces one response
  object reporting success or failure.

The response format is OPTIONAL for write operations. Implementations MAY
silently apply operations without producing response output. However,
implementations that support INSPECT MUST produce JSONL responses.

### 14.2. JSONL Framing

Each response line is a complete JSON object. Responses are streamed one
per line, enabling incremental processing by callers.

```
{"op":"inspect","address":"sheet[\"Q1 Budget\"]", ...}\n
{"op":"inspect","address":"sheet[\"Q1 Budget\"]/D2", ...}\n
```

For write operations:

```
{"op":"set","address":"sheet[\"Q1 Budget\"]/D7","status":"ok"}\n
{"op":"format","address":"sheet[\"Q1 Budget\"]/D7","status":"ok"}\n
{"op":"comment","address":"sheet[\"Q1 Budget\"]/D2","status":"ok"}\n
```

### 14.3. Response Object Schema (CDDL)

The response schema is defined using CDDL ([RFC 8610]).

```cddl
; Root response — one per JSONL line
response = inspect-response / operation-response

; ============================================================
; INSPECT Response
; ============================================================

inspect-response = {
  op:       "inspect"
  address:  tstr                      ; the original address expression
  matched:  uint                      ; number of elements matched
  elements: [* element]               ; matched elements
  ? error:  tstr                      ; present if address resolution failed
}

element = {
  type:         element-type
  ? index:      uint                  ; 1-based position in parent container
  ? of:         uint                  ; total siblings in parent container
  * identity                          ; type-specific identity fields
  ? content:    content-info          ; present when INCLUDE content
  ? properties: properties-info       ; present when INCLUDE properties
  ? children:   [* element]           ; present when DEPTH > 0
  ? comments:   [* comment-info]      ; present when element has comments
  ? context:    context-info          ; present when CONTEXT > 0
}

element-type = "heading" / "paragraph" / "run" / "table" / "row"
             / "cell" / "list" / "item" / "image" / "section"
             / "bookmark" / "content-control"             ; Word
             / "sheet" / "excel-row" / "excel-cell"       ; Excel
             / "slide" / "shape"                          ; PowerPoint

; Type-specific identity fields (always present in addressing layer)
identity = (
  ? level:       uint                 ; heading level (1-9)
  ? style:       tstr                 ; applied style name
  ? name:        tstr                 ; sheet name, shape name, bookmark name
  ? reference:   tstr                 ; cell reference (e.g., "D2")
  ? sheet:       tstr                 ; parent sheet name (for cells)
  ? placeholder: tstr                 ; "title" / "subtitle" / "body" / "notes"
  ? tag:         tstr                 ; content control tag
)

; Content layer — present when INCLUDE content is specified
content-info = {
  ? text:      tstr                   ; display text of the element
  ? value:     tstr                   ; raw value (cells — may differ from text)
  ? dataType:  tstr                   ; "string" / "number" / "boolean" / "date"
  ? cells:     [* tstr]              ; cell values for a row (shorthand)
}

; Properties layer — present when INCLUDE properties is specified
; Uses the same property names as the FORMAT operation (§7)
properties-info = {
  * tstr => any
}

; Comment information — always included when comments exist on the element
comment-info = {
  author:  tstr
  text:    tstr
  ? date:  tstr                       ; ISO 8601 datetime
}

; Context — sibling elements before and after the match
context-info = {
  before:  [* element]
  after:   [* element]
}

; ============================================================
; Operation Response
; ============================================================

operation-response = {
  op:        operation-type
  address:   tstr                     ; the AT address
  status:    "ok" / "error" / "warning"
  ? message: tstr                     ; human-readable detail
  ? detail:  any                      ; operation-specific result data
}

operation-type = "set" / "replace" / "insert-before" / "insert-after"
               / "delete" / "append" / "prepend" / "format" / "style"
               / "comment" / "insert-row" / "insert-column"
               / "merge-cells" / "duplicate" / "property"
```

### 14.4. INSPECT Response Examples

#### Addressing only — Excel sheet discovery

Request:
```
OFFICETALK/1.0
DOCTYPE excel

INSPECT sheet[1]
```

Response:
```jsonl
{"op":"inspect","address":"sheet[1]","matched":1,"elements":[{"type":"sheet","index":1,"of":2,"name":"Q1 Budget"}]}
```

#### Content with depth — Excel sheet rows

Request:
```
OFFICETALK/1.0
DOCTYPE excel

INSPECT sheet["Q1 Budget"]
  DEPTH 1
  INCLUDE content
```

Response:
```jsonl
{"op":"inspect","address":"sheet[\"Q1 Budget\"]","matched":1,"elements":[{"type":"sheet","name":"Q1 Budget","children":[{"type":"excel-row","index":1,"content":{"cells":["Department","Q1 Budget","Q1 Actual","Variance"]}},{"type":"excel-row","index":2,"content":{"cells":["Engineering","150000","162000","-12000"]}},{"type":"excel-row","index":3,"content":{"cells":["Marketing","80000","75000","5000"]}},{"type":"excel-row","index":4,"content":{"cells":["Sales","120000","118000","2000"]}},{"type":"excel-row","index":5,"content":{"cells":["Operations","95000","101000","-6000"]}},{"type":"excel-row","index":6,"content":{"cells":["HR","45000","43000","2000"]}}]}]}
```

#### Content and properties — Excel cell with comment

Request:
```
OFFICETALK/1.0
DOCTYPE excel

INSPECT sheet["Q1 Budget"]/D2
  INCLUDE content, properties
```

Response:
```jsonl
{"op":"inspect","address":"sheet[\"Q1 Budget\"]/D2","matched":1,"elements":[{"type":"excel-cell","reference":"D2","sheet":"Q1 Budget","index":4,"of":4,"content":{"value":"-12000","dataType":"string"},"properties":{"bold":false,"font-name":"Calibri","font-size":"11pt","number-format":"General"},"comments":[{"author":"OfficeTalk","text":"Engineering is $12K over budget. Verify contractor spend."}]}]}
```

#### Content with context — cell neighbors

Request:
```
OFFICETALK/1.0
DOCTYPE excel

INSPECT sheet["Q1 Budget"]/D2
  INCLUDE content
  CONTEXT 1
```

Response:
```jsonl
{"op":"inspect","address":"sheet[\"Q1 Budget\"]/D2","matched":1,"elements":[{"type":"excel-cell","reference":"D2","sheet":"Q1 Budget","content":{"value":"-12000"},"context":{"before":[{"type":"excel-cell","reference":"D1","content":{"value":"Variance"}}],"after":[{"type":"excel-cell","reference":"D3","content":{"value":"5000"}}]}}]}
```

#### Word document outline — all headings

Request:
```
OFFICETALK/1.0
DOCTYPE word

INSPECT body/heading
  INCLUDE content
```

Response:
```jsonl
{"op":"inspect","address":"body/heading","matched":3,"elements":[{"type":"heading","level":1,"index":1,"of":11,"style":"Heading1","content":{"text":"Introduction"}},{"type":"heading","level":1,"index":4,"of":11,"style":"Heading1","content":{"text":"Specifications"}},{"type":"heading","level":1,"index":7,"of":11,"style":"Heading1","content":{"text":"Getting Started"}}]}
```

#### Word table with structure

Request:
```
OFFICETALK/1.0
DOCTYPE word

INSPECT body/table[1]
  DEPTH 2
  INCLUDE content
```

Response:
```jsonl
{"op":"inspect","address":"body/table[1]","matched":1,"elements":[{"type":"table","index":8,"of":11,"children":[{"type":"row","index":1,"children":[{"type":"cell","index":1,"content":{"text":"Format"}},{"type":"cell","index":2,"content":{"text":"Extension"}},{"type":"cell","index":3,"content":{"text":"Status"}}]},{"type":"row","index":2,"children":[{"type":"cell","index":1,"content":{"text":"Word"}},{"type":"cell","index":2,"content":{"text":".docx"}},{"type":"cell","index":3,"content":{"text":"Supported"}}]}]}]}
```

#### PowerPoint slide overview

Request:
```
OFFICETALK/1.0
DOCTYPE powerpoint

INSPECT slide
  DEPTH 1
  INCLUDE content
```

Response:
```jsonl
{"op":"inspect","address":"slide","matched":4,"elements":[{"type":"slide","index":1,"of":4,"children":[{"type":"shape","placeholder":"title","content":{"text":"Q1 Business Review"}},{"type":"shape","placeholder":"subtitle","content":{"text":"Prepared by: Finance Team\nDate: March 2026"}}]},{"type":"slide","index":2,"of":4,"children":[{"type":"shape","placeholder":"title","content":{"text":"Revenue Highlights"}},{"type":"shape","placeholder":"body","content":{"text":"Total revenue: $2.4M (+12% YoY)\nNew customers: 47\nChurn rate: 2.1%"}}]},{"type":"slide","index":3,"of":4,"children":[{"type":"shape","placeholder":"title","content":{"text":"Key Risks"}},{"type":"shape","placeholder":"body","content":{"text":"Supply chain delays in APAC\nRegulatory changes in EU market\nTalent retention in Engineering"}}],"comments":[{"author":"OfficeTalk","text":"These risks need mitigation plans before the board meeting."}]},{"type":"slide","index":4,"of":4,"children":[{"type":"shape","placeholder":"title","content":{"text":"Next Steps"}},{"type":"shape","placeholder":"body","content":{"text":"1. Finalize Q2 budget allocations\n2. Launch customer retention program\n3. Hire 5 senior engineers\n4. Schedule board review for April 15"}}]}]}
```

### 14.5. Operation Response Examples

Write operations optionally produce response lines reporting the outcome of
each operation.

Request:
```
OFFICETALK/1.0
DOCTYPE excel

AT sheet["Q1 Budget"]/D7
  SET "-9000"
  FORMAT bold=true, border-bottom=medium

AT sheet["Q1 Budget"]/D2
  COMMENT "Verify contractor spend."
```

Response:
```jsonl
{"op":"set","address":"sheet[\"Q1 Budget\"]/D7","status":"ok"}
{"op":"format","address":"sheet[\"Q1 Budget\"]/D7","status":"ok"}
{"op":"comment","address":"sheet[\"Q1 Budget\"]/D2","status":"ok"}
```

Error example:
```jsonl
{"op":"set","address":"sheet[\"Missing\"]/A1","status":"error","message":"Address resolution failed: no sheet named 'Missing'."}
```

### 14.6. Content Type

The JSONL response format is identified by the media type:

```
application/officetalk-response+jsonl
```

Implementations that accept OfficeTalk documents and produce responses
SHOULD use this media type in Content-Type headers.

---

## ABNF Grammar Additions

The following additions extend the formal grammar in §13.

```abnf
; --- Inspect Operations ---
block            =/ inspect-block

inspect-block    = "INSPECT" SP address LF
                   *( inspect-modifier / comment-line )

inspect-modifier = depth-modifier / include-modifier / context-modifier

depth-modifier   = indent "DEPTH" SP 1*DIGIT LF
include-modifier = indent "INCLUDE" SP layer *( "," SP layer ) LF
context-modifier = indent "CONTEXT" SP 1*DIGIT LF

layer            = "content" / "properties"

indent           = 2*WSP
```

---

## Conformance

An implementation MAY support INSPECT operations, write operations, or both.
Implementations MUST declare which operation classes they support.

An implementation that supports INSPECT:
- MUST produce JSONL responses conforming to the CDDL schema in §14.3.
- MUST support the DEPTH, INCLUDE, and CONTEXT modifiers.
- MUST reject documents that mix INSPECT and write operations.

An implementation that supports write operations:
- MAY produce JSONL operation responses.
- If it produces operation responses, they MUST conform to the CDDL schema.

---

## References

- [RFC 5234] Augmented BNF for Syntax Specifications: ABNF
- [RFC 8610] Concise Data Definition Language (CDDL)
- [RFC 7464] JavaScript Object Notation (JSON) Text Sequences
