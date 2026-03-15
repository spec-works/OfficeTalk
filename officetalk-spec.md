# The `application/officetalk` Media Type

**Draft Specification — Version 1.0**

**Author:** Darrel Miller

**Date:** March 13, 2026

## Abstract

This document defines the `application/officetalk` media type, a structured
document format for expressing deterministic modifications to Microsoft Office
documents (Word, Excel, PowerPoint). OfficeTalk provides a human-readable,
LLM-friendly grammar for addressing content within Office documents and
specifying precise operations to transform that content.

OfficeTalk documents can be produced by any source: large language models,
command-line tools, template engines, scripted pipelines, or hand-authored by
developers. They are executed by an implementation library built on the Office
Open XML SDK. All operations are deterministic: given the same OfficeTalk
document and the same input Office document, the result is always identical.

## Table of Contents

1. [Introduction](#1-introduction)
2. [Media Type Definition](#2-media-type-definition)
3. [Document Structure](#3-document-structure)
4. [Addressing](#4-addressing)
5. [Operations](#5-operations)
6. [Data Types](#6-data-types)
7. [Formatting Properties](#7-formatting-properties)
8. [Processing Model](#8-processing-model)
9. [Validation](#9-validation)
10. [Security Considerations](#10-security-considerations)
11. [IANA Considerations](#11-iana-considerations)
12. [Examples](#12-examples)
13. [Formal Grammar](#13-formal-grammar)

---

## 1. Introduction

### 1.1. Purpose

OfficeTalk bridges the gap between natural language intent expressed by a user
to a large language model and the precise, low-level operations required to
modify an Office Open XML document. It serves as a portable, validatable
intermediate representation.

A producer (LLM, CLI tool, template engine, or application code) generates an
OfficeTalk document. A runtime library parses and validates it, resolves its
addresses against a target Office document, and applies the specified
operations using the Office Open XML SDK.

### 1.2. Design Principles

The format is governed by the following design principles:

1. **Generability** — The grammar uses line-oriented syntax, UPPERCASE
   keywords, simple quoting rules, and flat structure. These characteristics
   make it straightforward to produce from LLMs, CLI tools, template engines,
   and application code alike.

2. **Determinism** — Every operation has unambiguous semantics. No operation
   requires inference, content generation, or external data. The same
   OfficeTalk document applied to the same Office document always produces
   the same result.

3. **Validatability** — An OfficeTalk document can be syntactically parsed and
   semantically validated against a target document without invoking an LLM.
   Errors are reported with precise locations and categories.

4. **Addressability** — Content within Office documents is addressed using a
   simplified path syntax that avoids Office Open XML namespaces and deep
   nesting. The implementation resolves these addresses to concrete XML nodes.

5. **Composability** — An OfficeTalk document is a sequence of independent
   operation blocks. Blocks can be reordered, filtered, or composed from
   multiple sources.

### 1.3. Terminology

| Term | Definition |
|------|-----------|
| **OfficeTalk document** | A UTF-8 text document conforming to this specification |
| **Target document** | The Office Open XML document to be modified |
| **Operation block** | An address line followed by one or more operation lines |
| **Address** | A path expression identifying content in the target document |
| **Predicate** | A bracketed filter expression within an address segment |
| **Content block** | A multi-line text literal delimited by `<<<` and `>>>` |
| **Snapshot semantics** | All addresses are resolved against the original target document before any operations are applied |

### 1.4. Conventions

The key words "MUST", "MUST NOT", "REQUIRED", "SHALL", "SHALL NOT", "SHOULD",
"SHOULD NOT", "RECOMMENDED", "MAY", and "OPTIONAL" in this document are to be
interpreted as described in [RFC 2119].

---

## 2. Media Type Definition

| Field | Value |
|-------|-------|
| Type name | `application` |
| Subtype name | `officetalk` |
| File extension | `.otk` |
| Encoding | UTF-8 |
| Fragment identifier | OfficeTalk address (see [Section 4](#4-addressing)) |

---

## 3. Document Structure

### 3.1. Header

Every OfficeTalk document MUST begin with a version line followed by a
document type line:

```
OFFICETALK/1.0
DOCTYPE word
```

The version line declares the specification version. This document defines
version `1.0`.

The `DOCTYPE` line declares the target document type. Valid values are:

| Value | Target Format |
|-------|--------------|
| `word` | Word documents (.docx) |
| `excel` | Excel workbooks (.xlsx) |
| `powerpoint` | PowerPoint presentations (.pptx) |

### 3.2. Operation Blocks

The body of an OfficeTalk document consists of one or more operation blocks,
separated by blank lines. Each block begins with an `AT` line specifying the
target address, followed by one or more operation lines:

```
AT <address>
<OPERATION> [arguments]
<OPERATION> [arguments]
```

By default, an `AT` address MUST resolve to exactly one element. If the
address matches multiple elements, the implementation MUST report an
ambiguity error.

To explicitly target all matching elements, use the `EACH` modifier:

```
AT EACH <address>
<OPERATION> [arguments]
```

When `EACH` is used, the operations are applied to every element matching
the address. If no elements match, the implementation MUST report an error.
`EACH` is intended for bulk formatting, styling, and replacement operations.
Producers SHOULD exercise caution when combining `EACH` with destructive
operations like `DELETE`.

Operations within a block are applied sequentially to the same target. When a
content-producing operation (SET, INSERT BEFORE, INSERT AFTER) is followed by
a FORMAT or STYLE operation in the same block, the formatting applies to the
content just produced.

### 3.3. Comments

Lines beginning with `#` are comments and MUST be ignored by parsers.
Comments MAY appear anywhere except within content blocks.

```
# This is a comment
AT body/paragraph[1]
# Replace the greeting
REPLACE "Hello" WITH "Welcome"
```

### 3.4. Whitespace and Line Endings

- Lines are terminated by `LF` (U+000A) or `CRLF` (U+000D U+000A).
- Leading and trailing whitespace on operation lines is ignored.
- Blank lines (containing only whitespace) separate operation blocks.
- Indentation is not significant except within content blocks.

---

## 4. Addressing

### 4.1. Address Syntax

An address is a `/`-separated path of segments, each optionally followed by
one or more predicate expressions in square brackets:

```
segment[predicate]/segment[predicate]/...
```

Addresses are resolved relative to the target document's root. They do not
use Office Open XML namespace prefixes.

### 4.2. Path Segments

A path segment is a lowercase identifier naming a structural element:

| Segment | Meaning |
|---------|---------|
| `body` | Document body (Word) |
| `paragraph` | A paragraph |
| `heading` | A paragraph with a heading style |
| `run` | A text run within a paragraph |
| `table` | A table |
| `row` | A table row |
| `cell` | A table cell |
| `list` | A list |
| `item` | A list item |
| `image` | An inline or anchored image |
| `section` | A document section (Word) |
| `header` | Page header (Word) |
| `footer` | Page footer (Word) |
| `bookmark` | A named bookmark (Word) |
| `content-control` | A structured document tag / content control (Word) |
| `sheet` | A worksheet (Excel) |
| `range` | A cell range (Excel) |
| `column` | A column (Excel) |
| `slide` | A slide (PowerPoint) |
| `shape` | A shape on a slide (PowerPoint) |
| `title` | The title placeholder (PowerPoint) |
| `subtitle` | The subtitle placeholder (PowerPoint) |
| `notes` | Speaker notes (PowerPoint) |

### 4.3. Predicates

Predicates filter elements within a path segment. Multiple predicates on the
same segment are combined with AND semantics.

#### 4.3.1. Positional Predicate

A bare integer selects by 1-based position:

```
paragraph[3]         # third paragraph
table[1]/row[2]      # second row of first table
slide[5]             # fifth slide
```

#### 4.3.2. Key-Value Predicates

Named predicates filter by attribute:

```
heading[level=2]                        # h2 heading (must be unique, or use EACH)
heading[level=2, text="Background"]     # specific h2 by text
table[caption="Quarterly Results"]      # table by caption
content-control[tag="author-name"]      # content control by tag
sheet["Revenue"]                        # sheet by name (shorthand)
shape[name="Logo"]                      # shape by name
header[type=default]                    # default page header
```

A bare string predicate (without a key) is shorthand for the element's
primary identifier:

```
sheet["Revenue"]       ≡  sheet[name="Revenue"]
bookmark["intro"]      ≡  bookmark[name="intro"]
```

#### 4.3.3. Text Matching Predicates

Text content can be matched using several operators:

| Operator | Meaning | Example |
|----------|---------|---------|
| `text="..."` | Exact match | `paragraph[text="Hello World"]` |
| `text~="..."` | I-Regexp match ([RFC 9485]) | `paragraph[text~="^Chapter \d+"]` |
| `text^="..."` | Starts with | `paragraph[text^="In conclusion"]` |
| `text$="..."` | Ends with | `paragraph[text$="respectively."]` |
| `text*="..."` | Contains | `paragraph[text*="important"]` |

When a text predicate matches multiple elements, a positional predicate MAY
be appended to disambiguate:

```
paragraph[text*="revenue"][2]    # second paragraph containing "revenue"
```

#### 4.3.4. Heading-Scoped Addressing

When a `heading` segment appears in a non-terminal position (i.e., followed by
further path segments), it defines a **section scope**. The section scope
contains all sibling elements between the matched heading and the next heading
at the same or higher level (lower or equal outline level number).

```
body/heading[text="Methods"]/paragraph[1]
body/heading[level=2, text="Results"]/table[1]
body/heading[level=1]/paragraph[text*="summary"]
```

The first example addresses the first non-heading paragraph following the
"Methods" heading, up to (but not including) the next heading of equal or
higher level. The second example addresses the first table within the
"Results" section.

If the heading is the last heading at its level in the document, the section
scope extends to the end of the document body (or to the next heading at a
higher level).

Heading-scoped addressing is particularly useful for LLM-generated operations
where the human description refers to content by its section context (e.g.,
"change the paragraph under Operations") rather than by absolute position or
text content.

**Nesting:** Heading scopes may be nested to address content within
subsections:

```
body/heading[level=1, text="Chapter 3"]/heading[level=2, text="Analysis"]/paragraph[1]
```

This addresses the first paragraph within the "Analysis" subsection of
"Chapter 3".

#### 4.3.5. Excel Cell References

Excel addresses support standard cell reference notation as shorthand:

```
sheet["Sales"]/A1               ≡  sheet["Sales"]/cell[ref="A1"]
sheet["Sales"]/A1:D10           ≡  sheet["Sales"]/range[ref="A1:D10"]
sheet["Sales"]/column/B         ≡  sheet["Sales"]/column[ref="B"]
```

### 4.4. Word Document Addresses

```
body/paragraph[1]
body/paragraph[text*="conclusion"]
body/heading[level=1]
body/heading[level=2, text="Methods"]
body/heading[text="Introduction"]/paragraph[1]
body/heading[level=2, text="Results"]/table[1]
body/heading[level=1, text="Chapter 3"]/heading[level=2, text="Analysis"]/paragraph[1]
body/table[1]/row[3]/cell[2]
body/table[caption="Results"]/row[1]
body/list[1]/item[3]
body/image[1]
body/image[alt="Company Logo"]
body/section[2]
body/bookmark["references"]
body/content-control[tag="abstract"]
header[type=default]
header[type=first]
footer[type=default]
```

### 4.5. Excel Document Addresses

```
sheet["Revenue"]
sheet["Revenue"]/A1
sheet["Revenue"]/A1:D10
sheet["Revenue"]/row[5]
sheet["Revenue"]/column/C
sheet[1]/table["SalesData"]
sheet[1]/table["SalesData"]/row[2]/cell[3]
```

### 4.6. PowerPoint Document Addresses

```
slide[1]
slide[1]/title
slide[1]/subtitle
slide[1]/body
slide[3]/shape[name="Chart1"]
slide[3]/table[1]/row[1]/cell[2]
slide[2]/notes
slide[4]/image[1]
```

### 4.7. Address Resolution

Addresses are resolved against the target document as follows:

1. The implementation traverses the document structure left to right through
   each path segment.
2. At each segment, the implementation collects all matching elements of that
   type within the current scope.
3. Predicates filter the matching set. Positional predicates select by index;
   key-value predicates filter by attribute; text predicates filter by content.
4. **Heading scope:** When a `heading` segment is followed by additional path
   segments, the implementation MUST compute the section scope for each matched
   heading. The scope includes all sibling elements after the heading up to
   (but not including) the next heading at the same or higher level. Subsequent
   segments are resolved within this scope rather than within the heading
   element itself.
5. If the result set is empty, the address fails to resolve and the
   implementation MUST report an error.
6. If `EACH` is specified, the full result set is returned as the target.
7. If `EACH` is not specified and the result set contains more than one
   element, the implementation MUST report an ambiguity error.

---

## 5. Operations

### 5.1. Content Operations

#### 5.1.1. SET

Replaces the entire content of the addressed element.

```
AT body/paragraph[3]
SET "New paragraph text."

AT sheet["Sales"]/B7
SET "1250.00"
```

SET with a content block:

```
AT body/paragraph[3]
SET <<<
This paragraph now contains
multiple lines of text.
>>>
```

#### 5.1.2. REPLACE

Finds and replaces text within the addressed element. The first argument is
the search text; the second (after `WITH`) is the replacement.

```
AT body/paragraph[1]
REPLACE "FY2024" WITH "FY2025"
```

REPLACE operates on the text content of the addressed element. If the search
text spans multiple runs in the underlying XML, the implementation MUST
handle the cross-run replacement. If the search text is not found, the
implementation MUST report an error.

An optional `ALL` modifier replaces all occurrences within the element:

```
AT body
REPLACE ALL "colour" WITH "color"
```

#### 5.1.3. INSERT BEFORE / INSERT AFTER

Inserts new content before or after the addressed element. The inserted
content becomes a sibling of the target at the same structural level.

```
AT body/heading[text="Conclusion"]
INSERT BEFORE <<<
This new section appears just before the Conclusion.
>>>

AT body/paragraph[1]
INSERT AFTER "A new paragraph after the first one."
```

#### 5.1.4. DELETE

Removes the addressed element and its content.

```
AT body/paragraph[text*="DRAFT"]
DELETE

AT body/table[2]/row[5]
DELETE

AT slide[3]
DELETE
```

#### 5.1.5. APPEND / PREPEND

Adds content to the beginning or end of the addressed element's existing
content, without replacing it.

```
AT body/paragraph[1]
APPEND " (see Appendix A)"

AT sheet["Notes"]/A1
PREPEND "UPDATED: "
```

### 5.2. Formatting Operations

#### 5.2.1. FORMAT

Applies formatting properties to the addressed element. Properties are
specified as comma-separated `key=value` pairs.

```
AT body/paragraph[1]
FORMAT bold=true, font-size=14pt, color=#2B579A

AT sheet["Revenue"]/A1:A10
FORMAT font-name="Calibri", fill-color=#F2F2F2, number-format="#,##0.00"
```

Multiple FORMAT lines in the same block are cumulative:

```
AT body/heading[level=1]
FORMAT font-name="Aptos Display", font-size=28pt
FORMAT color=#1F3864, spacing-after=12pt
```

When FORMAT follows a content-producing operation (SET, INSERT BEFORE,
INSERT AFTER, APPEND, PREPEND) in the same block, it applies to the
produced content:

```
AT body/heading[text="Summary"]
INSERT AFTER "Key Findings"
FORMAT bold=true, font-size=16pt, style="Heading 2"
```

#### 5.2.2. STYLE

Applies a named document style to the addressed element. The style MUST
exist in the target document or its attached template.

```
AT body/paragraph[5]
STYLE "Heading 2"

AT body/table[1]
STYLE "Grid Table 4 - Accent 1"
```

### 5.3. Structural Operations

#### 5.3.1. Table Operations

```
AT body/table[1]/row[3]
INSERT ROW AFTER

AT body/table[1]/row[3]
INSERT ROW BEFORE

AT body/table[1]/row[5]
DELETE ROW

AT body/table[1]/cell[1]
INSERT COLUMN AFTER

AT body/table[1]/cell[1]
INSERT COLUMN BEFORE

AT body/table[1]/cell[3]
DELETE COLUMN

AT body/table[1]/row[2]/cell[1]
MERGE CELLS TO row[2]/cell[3]
```

When INSERT ROW is followed by SET operations, the SET values populate cells
in the newly inserted row by position:

```
AT body/table[1]/row[3]
INSERT ROW AFTER
SET CELLS "Product D", "450", "12%", "Active"
```

#### 5.3.2. Slide Operations (PowerPoint)

```
AT slide[2]
INSERT SLIDE AFTER
SET title "New Section"
SET subtitle "Overview of Changes"

AT slide[5]
DELETE SLIDE

AT slide[3]
DUPLICATE SLIDE
```

#### 5.3.3. Sheet Operations (Excel)

```
ADD SHEET "Summary"

AT sheet["Old Name"]
RENAME SHEET "New Name"

AT sheet["Temp"]
DELETE SHEET
```

Sheet operations do not require an `AT` line when the operation specifies
its target inline (e.g., `ADD SHEET`).

### 5.4. Metadata Operations

#### 5.4.1. PROPERTY

Sets document-level properties.

```
PROPERTY title="Quarterly Report Q1 2026"
PROPERTY author="Finance Team"
PROPERTY subject="Financial review"
```

PROPERTY lines do not require an `AT` address. They apply to the document
as a whole.

---

## 6. Data Types

### 6.1. Strings

Strings are enclosed in double quotes. The following escape sequences are
supported:

| Sequence | Meaning |
|----------|---------|
| `\"` | Literal double quote |
| `\\` | Literal backslash |
| `\n` | Newline (U+000A) |
| `\t` | Tab (U+0009) |

```
"Hello, World!"
"She said \"hello\" to them."
"Line one\nLine two"
```

### 6.2. Numbers

Decimal integers and floating-point numbers. Negative values use a leading
minus sign.

```
42
3.14
-7
0.5
```

### 6.3. Booleans

The literal values `true` and `false` (case-sensitive).

### 6.4. Colors

Colors are specified as hexadecimal RGB or RRGGBB values prefixed with `#`,
or as named CSS colors:

```
#2B579A
#FF0000
#00FF0080       # with alpha channel
red
cornflowerblue
```

### 6.5. Lengths

Numeric values with a unit suffix:

| Unit | Meaning |
|------|---------|
| `pt` | Points (1/72 inch) |
| `in` | Inches |
| `cm` | Centimeters |
| `mm` | Millimeters |
| `px` | Pixels (at 96 DPI) |
| `%` | Percentage of parent |
| `emu` | English Metric Units (Office native) |

```
12pt
1.5in
2.54cm
50%
914400emu
```

### 6.6. Content Blocks

Multi-line text content is enclosed between `<<<` and `>>>` delimiters,
each on its own line:

```
<<<
First paragraph of inserted text.

Second paragraph, separated by a blank line.
This is still part of the second paragraph.
>>>
```

Rules for content blocks:

- The `<<<` delimiter MUST appear on its own line (after the operation keyword
  and any required whitespace).
- The `>>>` delimiter MUST appear on its own line.
- Content between delimiters is literal text. No escape processing is
  performed.
- Blank lines within a content block create paragraph breaks.
- Leading and trailing blank lines within the block are stripped.

---

## 7. Formatting Properties

### 7.1. Text (Run) Properties

| Property | Type | Description |
|----------|------|-------------|
| `font-name` | string | Font family name |
| `font-size` | length | Font size |
| `bold` | boolean | Bold weight |
| `italic` | boolean | Italic style |
| `underline` | enum | Underline style: `single`, `double`, `dotted`, `dashed`, `wavy`, `none` |
| `strikethrough` | boolean | Strikethrough |
| `color` | color | Text color |
| `highlight` | color | Text highlight color |
| `superscript` | boolean | Superscript position |
| `subscript` | boolean | Subscript position |
| `small-caps` | boolean | Small capitals |
| `all-caps` | boolean | All capitals |

### 7.2. Paragraph Properties

| Property | Type | Description |
|----------|------|-------------|
| `alignment` | enum | `left`, `center`, `right`, `justify` |
| `spacing-before` | length | Space before paragraph |
| `spacing-after` | length | Space after paragraph |
| `line-spacing` | number/length | Line spacing (number = multiple, length = exact) |
| `indent-left` | length | Left indent |
| `indent-right` | length | Right indent |
| `indent-first-line` | length | First line indent |
| `indent-hanging` | length | Hanging indent |
| `keep-with-next` | boolean | Keep with next paragraph |
| `page-break-before` | boolean | Page break before paragraph |
| `outline-level` | number | Outline level (0-8) |

### 7.3. Table Properties

| Property | Type | Description |
|----------|------|-------------|
| `width` | length | Table width |
| `alignment` | enum | `left`, `center`, `right` |
| `border` | enum | `all`, `outside`, `inside`, `none` |
| `border-color` | color | Border color |
| `border-width` | length | Border width |
| `cell-padding` | length | Default cell padding |

### 7.4. Cell Properties

| Property | Type | Description |
|----------|------|-------------|
| `fill-color` | color | Cell background / fill color |
| `vertical-align` | enum | `top`, `center`, `bottom` |
| `border` | enum | `all`, `none`, or per-side |
| `width` | length | Cell / column width |
| `wrap-text` | boolean | Wrap text in cell (Excel) |
| `number-format` | string | Number format string (Excel) |

### 7.5. Image Properties

| Property | Type | Description |
|----------|------|-------------|
| `width` | length | Image width |
| `height` | length | Image height |
| `alt` | string | Alt text |
| `position` | enum | `inline`, `anchor` (Word) |

---

## 8. Processing Model

### 8.1. Overview

Processing an OfficeTalk document against a target Office document proceeds
in three phases: resolution, validation, and execution.

### 8.2. Resolution Phase

The implementation MUST resolve all addresses in all operation blocks against
the target document **before** any operations are applied. This is the
**snapshot semantics** rule.

Rationale: Operations may insert or delete elements, changing positional
indices. Resolving all addresses against the original document ensures that
`paragraph[5]` always refers to the fifth paragraph in the original, not a
shifted position after prior operations.

The resolved addresses are bound to the specific XML nodes they identify.

### 8.3. Validation Phase

After resolution, the implementation performs semantic validation:

1. **Address validity** — All addresses resolved to exactly one element
   (or an appropriate set for bulk operations like REPLACE ALL).
2. **Operation applicability** — Each operation is valid for its target
   element type (e.g., SET CELLS is only valid after INSERT ROW on a table).
3. **Type correctness** — All FORMAT property values are valid for their
   declared types.
4. **Style existence** — All STYLE references name styles present in the
   target document.
5. **Search text existence** — All REPLACE search strings exist in their
   target elements.

If validation fails, the implementation MUST NOT modify the target document
and MUST report all errors.

### 8.4. Execution Phase

Operations are applied in document order (top to bottom). Within an operation
block, operations are applied sequentially.

The implementation SHOULD apply all operations atomically: either all
operations succeed and the document is modified, or no operations are applied
and the original document is unchanged. Implementations that cannot guarantee
atomicity MUST document this limitation.

### 8.5. Conflict Detection

The implementation SHOULD detect and report conflicts between operation
blocks that target the same or overlapping elements:

- Two SET operations on the same element (last write wins, but warn).
- A DELETE followed by a FORMAT on the same element (error).
- Overlapping MERGE CELLS operations (error).

---

## 9. Validation

### 9.1. Syntactic Validation

Syntactic validation ensures the OfficeTalk document conforms to the grammar
defined in [Section 13](#13-formal-grammar). This can be performed without
access to a target document.

Syntactic errors include:
- Missing or malformed header (version, DOCTYPE)
- Unrecognized operation keywords
- Malformed addresses (unbalanced brackets, invalid segments)
- Unterminated strings or content blocks
- Invalid data type literals (e.g., malformed color codes)

### 9.2. Semantic Validation

Semantic validation requires access to the target document. It verifies:

- Addresses resolve to existing elements
- Addresses are unambiguous (single-element resolution)
- Operations are applicable to their target element types
- Referenced styles exist
- REPLACE search strings are present in target content
- FORMAT property names and value types are valid for the target element
- Structural operations are valid (e.g., not deleting the only row in a table)

### 9.3. Error Categories

| Category | Description |
|----------|------------|
| `SYNTAX` | Grammar violation in the OfficeTalk document |
| `ADDRESS_NOT_FOUND` | Address does not match any element |
| `ADDRESS_AMBIGUOUS` | Address matches multiple elements |
| `INVALID_OPERATION` | Operation not applicable to target element type |
| `INVALID_VALUE` | Property value has wrong type or is out of range |
| `MISSING_STYLE` | Referenced style not found in document |
| `SEARCH_NOT_FOUND` | REPLACE search text not found in target |
| `CONFLICT` | Conflicting operations on same element |
| `STRUCTURAL` | Invalid structural operation |

### 9.4. Warnings

| Category | Description |
|----------|------------|
| `STYLE_OVERRIDE` | FORMAT overrides properties set by STYLE in same block |
| `REDUNDANT_OP` | Operation has no effect (e.g., SET to existing value) |
| `DEPRECATED` | Use of deprecated syntax or property name |

---

## 10. Security Considerations

1. **No code execution** — OfficeTalk documents MUST NOT contain executable
   code. Implementations MUST NOT evaluate expressions, execute scripts, or
   invoke external services during processing.

2. **No external references** — Content blocks contain literal text only.
   Implementations MUST NOT fetch external resources (URLs, file paths)
   referenced in OfficeTalk content.

3. **Input sanitization** — Implementations MUST validate all string inputs
   to prevent injection of malicious Office Open XML content.

4. **Resource limits** — Implementations SHOULD enforce limits on document
   size, number of operations, content block length, and regex complexity
   to prevent denial-of-service.

5. **Regex safety** — Text matching predicates using `text~="..."` use
   I-Regexp [RFC 9485], a restricted regular expression syntax designed for
   interoperability. The limited feature set of I-Regexp inherently mitigates
   ReDoS risks. Implementations SHOULD additionally enforce execution
   timeouts as a defense-in-depth measure.

---

## 11. IANA Considerations

### 11.1. Media Type Registration

```
Type name:               application
Subtype name:            officetalk
Required parameters:     none
Optional parameters:     version (default "1.0")
Encoding considerations: UTF-8
Security considerations: See Section 10
Published specification: This document
Applications:            Microsoft Office document transformation
Fragment identifier:     OfficeTalk address path (Section 4)
File extension:          .otk
```

---

## 12. Examples

### 12.1. Word: Basic Text Editing

```
OFFICETALK/1.0
DOCTYPE word

# Fix a typo in the introduction
AT body/paragraph[text*="teh company"]
REPLACE "teh" WITH "the"

# Update the document title
AT body/heading[level=1]
SET "Annual Report — FY2026"
FORMAT font-size=28pt, color=#1F3864

# Remove the draft watermark paragraph
AT body/paragraph[text="DRAFT — DO NOT DISTRIBUTE"]
DELETE
```

### 12.2. Word: Insert a New Section

```
OFFICETALK/1.0
DOCTYPE word

AT body/heading[text="Conclusion"]
INSERT BEFORE <<<
Recommendations

Based on the findings presented in this report, the committee
recommends the following actions for the upcoming fiscal year.

1. Increase investment in renewable energy infrastructure.
2. Expand the remote work pilot program to all departments.
3. Commission an independent audit of supply chain practices.
>>>
FORMAT style="Heading 1"
```

### 12.3. Word: Heading-Scoped Editing

```
OFFICETALK/1.0
DOCTYPE word

# Change the first paragraph under the Operations heading
AT body/heading[text="Operations"]/paragraph[1]
SET "The operations squad is responsible for all production deployments."

# Update a table within a specific section
AT body/heading[level=2, text="Financial Results"]/table[1]/row[2]/cell[3]
SET "1,250,000"

# Replace text only within a subsection
AT body/heading[level=1, text="Chapter 3"]/heading[level=2, text="Analysis"]/paragraph[text*="outdated"]
REPLACE "outdated methodology" WITH "revised methodology"
```

### 12.4. Word: Table Manipulation

```
OFFICETALK/1.0
DOCTYPE word

# Add a new row to the quarterly results table
AT body/table[caption="Quarterly Results"]/row[4]
INSERT ROW AFTER
SET CELLS "Q4", "$2.1M", "$1.8M", "16.7%"

# Highlight the header row
AT body/table[caption="Quarterly Results"]/row[1]
FORMAT bold=true, fill-color=#2B579A, color=#FFFFFF

# Widen the first column
AT body/table[caption="Quarterly Results"]/cell[1]
FORMAT width=2in
```

### 12.5. Excel: Update a Spreadsheet

```
OFFICETALK/1.0
DOCTYPE excel

# Update a cell value
AT sheet["Revenue"]/B7
SET "1250.00"
FORMAT number-format="#,##0.00", bold=true

# Format a header range
AT sheet["Revenue"]/A1:F1
FORMAT font-name="Calibri", font-size=12pt, bold=true
FORMAT fill-color=#4472C4, color=#FFFFFF

# Add a new row of data
AT sheet["Revenue"]/row[15]
INSERT ROW AFTER
SET CELLS "Product E", "North", "450", "12.5%", "Active", "2026-01-15"

# Rename a worksheet
AT sheet["Sheet1"]
RENAME SHEET "Dashboard"
```

### 12.6. PowerPoint: Update a Presentation

```
OFFICETALK/1.0
DOCTYPE powerpoint

# Update the title slide
AT slide[1]/title
SET "Q1 2026 Business Review"
FORMAT font-size=40pt, bold=true, color=#1F3864

AT slide[1]/subtitle
SET "Prepared by the Strategy Team — March 2026"

# Add a new slide after slide 3
AT slide[3]
INSERT SLIDE AFTER
SET title "Key Metrics"
SET body <<<
Revenue grew 12% year-over-year.

Customer retention improved to 94%.

Three new enterprise clients onboarded.
>>>

# Update speaker notes
AT slide[5]/notes
SET "Remember to mention the pending regulatory review."

# Delete the appendix slide
AT slide[text*="Appendix — Old Data"]
DELETE SLIDE
```

### 12.7. Word: Comprehensive Document Rewrite

```
OFFICETALK/1.0
DOCTYPE word

# Update document metadata
PROPERTY title="Project Phoenix — Technical Design Document"
PROPERTY author="Engineering Team"

# Restyle the main heading
AT body/heading[level=1]
SET "Project Phoenix"
FORMAT font-name="Aptos Display", font-size=32pt, color=#2B579A

# Replace outdated terminology throughout the document body
AT body
REPLACE ALL "legacy system" WITH "heritage platform"
REPLACE ALL "Phase 1" WITH "Phase One"

# Insert an abstract before the first paragraph
AT body/paragraph[1]
INSERT BEFORE <<<
Abstract

This document describes the technical architecture for Project Phoenix,
a next-generation platform for real-time data processing. It covers
system design, API specifications, deployment topology, and operational
runbooks.
>>>

# Format the abstract heading
AT body/paragraph[text="Abstract"]
STYLE "Heading 2"

# Add a row to the milestones table
AT body/table[caption="Milestones"]/row[6]
INSERT ROW AFTER
SET CELLS "M7", "Production Readiness Review", "2026-06-15", "Pending"
FORMAT fill-color=#FFF2CC
```

---

## 13. Formal Grammar

The following ABNF grammar (per [RFC 5234]) defines the syntax of an
OfficeTalk document. This grammar is simplified for clarity; implementations
SHOULD consult the normative prose in preceding sections for full semantics.

```abnf
document         = header 1*( block / blank-line / comment-line )

header           = version-line doctype-line
version-line     = "OFFICETALK/" version LF
version          = 1*DIGIT "." 1*DIGIT
doctype-line     = "DOCTYPE" SP doctype LF
doctype          = "word" / "excel" / "powerpoint"

block            = ( operation-block / property-block / sheet-block )
operation-block  = address-line 1*( operation-line / comment-line )
address-line     = "AT" [ SP "EACH" ] SP address LF
property-block   = "PROPERTY" SP key "=" value LF
sheet-block      = ("ADD SHEET" / "DELETE SHEET" / "RENAME SHEET") SP quoted-string LF

operation-line   = ( set-op / replace-op / insert-op / delete-op
                   / append-op / prepend-op / format-op / style-op
                   / structural-op ) LF

; --- Content Operations ---
set-op           = "SET" SP ( quoted-string / content-block / cells-clause )
cells-clause     = "CELLS" SP quoted-string *( "," SP quoted-string )
replace-op       = ["REPLACE" / "REPLACE ALL"] SP quoted-string SP "WITH" SP
                   ( quoted-string / content-block )
insert-op        = ( "INSERT BEFORE" / "INSERT AFTER" ) SP
                   ( quoted-string / content-block / slide-props )
delete-op        = "DELETE" [ SP ( "ROW" / "COLUMN" / "SLIDE" / "SHEET" ) ]
append-op        = "APPEND" SP ( quoted-string / content-block )
prepend-op       = "PREPEND" SP ( quoted-string / content-block )

; --- Formatting Operations ---
format-op        = "FORMAT" SP property-list
property-list    = property *( "," SP property )
property         = key "=" prop-value
prop-value       = quoted-string / number / boolean / color / length
style-op         = "STYLE" SP quoted-string

; --- Structural Operations ---
structural-op    = row-op / column-op / merge-op / slide-op
row-op           = "INSERT ROW" SP ( "BEFORE" / "AFTER" )
column-op        = "INSERT COLUMN" SP ( "BEFORE" / "AFTER" )
merge-op         = "MERGE CELLS TO" SP address
slide-op         = ( "INSERT SLIDE" SP ( "BEFORE" / "AFTER" ) )
                 / "DUPLICATE SLIDE"
slide-props      = *( set-op LF )

; --- Address ---
address          = segment *( "/" segment )
segment          = identifier *( "[" predicate "]" )
identifier       = 1*( ALPHA / "-" )
predicate        = positional / key-value-pred / bare-string / cell-ref
positional       = 1*DIGIT
key-value-pred   = pred-key pred-op pred-val *( "," SP pred-key pred-op pred-val )
pred-key         = 1*( ALPHA / "-" )
pred-op          = "=" / "~=" / "^=" / "$=" / "*="
pred-val         = quoted-string / number / boolean
bare-string      = quoted-string
cell-ref         = 1*ALPHA 1*DIGIT [ ":" 1*ALPHA 1*DIGIT ]

; --- Data Types ---
quoted-string    = DQUOTE *( escaped-char / safe-char ) DQUOTE
escaped-char     = "\" ( DQUOTE / "\" / "n" / "t" )
safe-char        = %x20-21 / %x23-5B / %x5D-7E / UTF8-tail
number           = [ "-" ] 1*DIGIT [ "." 1*DIGIT ]
boolean          = "true" / "false"
color            = "#" 6*8HEXDIG / color-name
color-name       = 1*ALPHA
length           = number unit
unit             = "pt" / "in" / "cm" / "mm" / "px" / "%" / "emu"

; --- Content Block ---
content-block    = "<<<" LF *content-line ">>>" LF
content-line     = *safe-char LF

; --- Whitespace ---
blank-line       = *WSP LF
comment-line     = "#" *safe-char LF
LF               = %x0A / ( %x0D %x0A )

; --- Core Rules (from RFC 5234) ---
ALPHA            = %x41-5A / %x61-7A
DIGIT            = %x30-39
HEXDIG           = DIGIT / "A" / "B" / "C" / "D" / "E" / "F"
                 / "a" / "b" / "c" / "d" / "e" / "f"
DQUOTE           = %x22
SP               = %x20
WSP              = SP / %x09
UTF8-tail        = %x80-BF
```

---

## Appendix A. Design Rationale

### A.1. Why Not JSON or YAML?

JSON and YAML were considered and rejected:

- **JSON**: Verbose, requires escaping in strings, deeply nested for
  structural operations, cumbersome multi-line strings. LLMs frequently
  produce malformed JSON (trailing commas, unescaped quotes).

- **YAML**: Whitespace-sensitive indentation is fragile for LLM generation.
  YAML's type coercion (e.g., `no` → `false`) causes subtle bugs.

The OfficeTalk grammar is purpose-built for the intersection of broad
producer support (LLMs, CLI tools, template engines, scripted pipelines)
and machine parseability.

### A.2. Why Not XPath for Addressing?

Office Open XML documents use deeply nested XML with verbose namespace
prefixes (e.g., `w:`, `a:`, `r:`). XPath expressions for OOXML are brittle,
hard for LLMs to generate correctly, and expose implementation details that
the OfficeTalk user should not need to know.

The OfficeTalk addressing scheme provides a semantic abstraction layer that
the implementation maps to concrete XML nodes internally. This benefits all
producers — LLMs avoid namespace hallucination, and CLI tools and template
engines avoid coupling to Open XML internals.

### A.3. Why Snapshot Semantics?

Snapshot semantics (resolving all addresses before applying any operations)
prevent a class of bugs where early operations shift positional indices,
causing later operations to target the wrong elements. This makes OfficeTalk
documents easier to produce correctly — whether by an LLM reasoning about the
current document state, or by a CLI tool constructing operations from
parameters — since all producers can reason about the document as it currently
exists rather than tracking mutations.

---

## Appendix B. Implementation Notes

### B.1. Office Open XML SDK Integration

An implementation library using the Open XML SDK (C#/.NET) should:

1. Open the target document as a read-write `WordprocessingDocument`,
   `SpreadsheetDocument`, or `PresentationDocument`.
2. Parse the OfficeTalk document into an AST.
3. Walk the AST, resolving each address to one or more Open XML elements.
4. Store resolved element references.
5. Validate all operations against their resolved targets.
6. Apply operations sequentially, using the pre-resolved references.
7. Save the modified document.

### B.2. Address-to-OpenXML Mapping (Word)

| OfficeTalk Address | Open XML Element |
|--------------------|-----------------|
| `body/paragraph[n]` | `Body.Elements<Paragraph>()[n-1]` |
| `body/heading[level=n]` | `Paragraph` where `ParagraphProperties.ParagraphStyleId` maps to heading level `n` |
| `body/heading[text="X"]/paragraph[1]` | First non-heading `Paragraph` after the matched heading, up to the next heading at the same or higher level |
| `body/table[n]/row[m]/cell[k]` | `Table[n-1].Elements<TableRow>()[m-1].Elements<TableCell>()[k-1]` |
| `body/bookmark["name"]` | `BookmarkStart` where `Name == "name"` |
| `header[type=default]` | `HeaderPart` referenced by `SectionProperties.HeaderReference` with `Type == Default` |

### B.3. Regex Engine Requirements

Text matching predicates (`text~="..."`) use I-Regexp syntax as defined in
[RFC 9485]. Implementations MUST map I-Regexp patterns to their platform's
native regex dialect following the procedures in RFC 9485, Section 5.
For .NET implementations, this maps to the ECMAScript-compatible subset
described in RFC 9485, Section 5.3.

---

## References

- [RFC 2119] Bradner, S., "Key words for use in RFCs to Indicate Requirement
  Levels", BCP 14, RFC 2119, March 1997.
- [RFC 5234] Crocker, D. and P. Overell, "Augmented BNF for Syntax
  Specifications: ABNF", STD 68, RFC 5234, January 2008.
- [RFC 6838] Freed, N., Klensin, J., and T. Hansen, "Media Type
  Specifications and Registration Procedures", BCP 13, RFC 6838,
  January 2013.
- [RFC 9485] Bormann, C. and T. Bray, "I-Regexp: An Interoperable
  Regular Expression Format", RFC 9485, October 2023.
- [ECMA-376] ECMA International, "Office Open XML File Formats", ECMA-376,
  5th Edition, December 2021.
- [Open XML SDK] Microsoft, "Open XML SDK Documentation",
  https://docs.microsoft.com/en-us/office/open-xml/open-xml-sdk
