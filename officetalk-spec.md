# The `application/officetalk` Media Type

**Draft Specification — Version 1.0**

**Author:** Darrel Miller

**Date:** March 13, 2026

## Abstract

This document defines the `application/officetalk` media type, a structured
document format for expressing deterministic operations on Microsoft Office
documents (Word, Excel, PowerPoint). OfficeTalk provides a human-readable,
LLM-friendly grammar for addressing content within Office documents and
specifying precise operations to read or transform that content.

OfficeTalk supports two classes of operation: **write operations** that modify
document content (SET, DELETE, FORMAT, etc.) and **inspect operations** that
read document structure and content without modification. A single OfficeTalk
document contains either write operations or inspect operations, never both.

OfficeTalk documents can be produced by any source: large language models,
command-line tools, template engines, scripted pipelines, or hand-authored by
developers. They are executed by an implementation library built on the Office
Open XML SDK. All write operations are deterministic: given the same OfficeTalk
document and the same input Office document, the result is always identical.
Inspect operations produce structured JSONL responses describing the target
document's content and structure.

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
14. [Response Format](#14-response-format)

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

### 3.2. Document Classification

An OfficeTalk document is either a **write document** or a **read document**:

- A **write document** contains operation blocks (§3.3) and optional property
  settings (§5.4). It modifies the target document.
- A **read document** contains inspect blocks (§3.4). It reads the target
  document without modification.

A single OfficeTalk document MUST NOT contain both write operations
(`AT` blocks, `PROPERTY` statements) and `INSPECT` operations. Implementations
MUST reject documents that mix the two.

### 3.3. Operation Blocks

The body of a write document consists of one or more operation blocks,
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

### 3.4. Inspect Blocks

The body of a read document consists of one or more inspect blocks,
separated by blank lines. Each block begins with an `INSPECT` line
specifying the target address, optionally followed by indented modifier
lines:

```
INSPECT <address>
  [DEPTH <integer>]
  [INCLUDE <layer> *("," <layer>)]
  [CONTEXT <integer>]
```

Multiple `INSPECT` blocks in a single document are processed sequentially.
Each produces one result object in the JSONL response (see
[Section 14](#14-response-format)).

Inspect blocks reuse the same address syntax as operation blocks
([Section 4](#4-addressing)). An `INSPECT` address MAY resolve to one or
multiple elements; unlike `AT`, no ambiguity error is raised when multiple
elements match.

### 3.5. Comments

Lines beginning with `#` are comments and MUST be ignored by parsers.
Comments MAY appear anywhere except within content blocks.

```
# This is a comment
AT body/paragraph[1]
# Replace the greeting
REPLACE "Hello" WITH "Welcome"
```

### 3.6. Whitespace and Line Endings

- Lines are terminated by `LF` (U+000A) or `CRLF` (U+000D U+000A).
- Leading and trailing whitespace on operation lines is ignored.
- Blank lines (containing only whitespace) separate operation blocks.
- Indentation is not significant except within content blocks and
  INSPECT modifier lines (DEPTH, INCLUDE, CONTEXT), which MUST be
  indented by at least two spaces relative to their INSPECT line.

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
| `run` | A text span within a paragraph (see §4.3.5) |
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

#### 4.3.5. Run (Text Span) Resolution

The `run` segment addresses a contiguous span of text within a paragraph,
not an underlying XML run element (`<w:r>`). When a `run` segment includes a
text predicate, the implementation MUST locate the matching text within the
parent scope and isolate it as a targetable range.

```
body/paragraph[1]/run[text="OfficeTalk"]      # the text "OfficeTalk" in paragraph 1
body/paragraph[3]/run[text*="important"]       # text containing "important"
```

If the matched text falls within a single structural run, that run (or a
portion of it) is the target. If the matched text spans part of a larger
run, the implementation MUST split the underlying run to isolate the matched
span. This ensures that operations like FORMAT apply only to the targeted
text, not the entire structural run.

**Rationale:** LLMs generating OfficeTalk do not have visibility into how a
document's text is divided into structural runs. A paragraph displayed as
"AI agents use OfficeTalk to edit documents." may be stored as one run or
many. The `run` segment provides a content-based way to target specific text
regardless of the underlying structure.

All text matching operators (`text=`, `text*=`, `text^=`, `text$=`,
`text~=`) are supported on `run` segments with the same semantics as on
other elements (see §4.3.3).

#### 4.3.6. Excel Cell References

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
body/paragraph[1]/run[text="OfficeTalk"]
body/heading[text="Methods"]/paragraph[1]/run[text*="conclusion"]
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

#### 5.1.6. INSERT IMAGE

Inserts an image before or after the addressed element. The image source
is specified as a file path or URL in a quoted string.

```
AT body/paragraph[3]
INSERT IMAGE AFTER "charts/revenue-q1.png"

AT body/heading[text="Results"]
INSERT IMAGE AFTER "https://example.com/diagram.png"
  alt="Architecture diagram"
  width=6in
  height=4in
```

Image properties are specified as `key=value` pairs on indented lines
following the INSERT IMAGE line:

| Property | Type   | Description |
|----------|--------|-------------|
| `alt`    | string | Alt text for accessibility |
| `width`  | length | Display width |
| `height` | length | Display height |
| `position` | enum | `inline` (default) or `anchor` (Word only) |

If only `width` or `height` is specified, the other dimension is computed
to maintain the original aspect ratio. If neither is specified, the image
is inserted at its native dimensions.

The image source MUST be a file path (relative to the working directory)
or an HTTPS URL. Implementations MUST resolve relative paths against the
working directory, not the OfficeTalk document location. Implementations
MUST NOT follow redirects to non-HTTPS URLs.

In Word documents, INSERT IMAGE BEFORE/AFTER inserts the image as a new
paragraph-level element (an inline drawing within its own paragraph). In
PowerPoint documents, the image is inserted as a new picture shape on the
addressed slide.

#### 5.1.7. INSERT TABLE

Creates a new table before or after the addressed element with the
specified dimensions.

```
AT body/heading[text="Results"]
INSERT TABLE AFTER rows=3, columns=4
```

The table is created with empty cells. Use subsequent SET CELLS operations
to populate the table content:

```
AT body/heading[text="Results"]
INSERT TABLE AFTER rows=3, columns=4
SET CELLS "Product", "Q1", "Q2", "Q3"

AT body/table[caption="Results"]/row[2]
SET CELLS "Widget A", "100", "120", "135"

AT body/table[caption="Results"]/row[3]
SET CELLS "Widget B", "200", "210", "225"
```

Table properties may be specified as indented `key=value` pairs:

| Property | Type   | Description |
|----------|--------|-------------|
| `rows`   | number | Number of rows (required) |
| `columns` | number | Number of columns (required) |
| `caption` | string | Table caption / alt text |
| `width`  | length | Table width (default: 100% of page width) |

The first SET CELLS operation in the same block populates the first row
of the newly inserted table. INSERT TABLE is valid in Word and PowerPoint
documents.

#### 5.1.8. LINK

Creates or replaces a hyperlink on the addressed text. The target URL is
a quoted string.

```
AT body/paragraph[1]/run[text="click here"]
LINK "https://example.com/report"

AT body/paragraph[3]/run[text="OfficeTalk specification"]
LINK "https://github.com/spec-works/OfficeTalk"
```

LINK creates a clickable hyperlink on the addressed run or element. When
applied to a paragraph, the entire paragraph text becomes a hyperlink.
When applied to a run, only that run's text is linked.

LINK can follow a SET or INSERT operation to create linked text in a single
block:

```
AT body/paragraph[3]
INSERT AFTER "See the full report."
LINK "https://example.com/report"
```

To remove a hyperlink, use FORMAT:

```
AT body/paragraph[1]/run[text="click here"]
FORMAT href=none
```

#### 5.1.9. INSERT LIST

Creates a new bulleted or numbered list before or after the addressed
element.

```
AT body/heading[text="Action Items"]
INSERT LIST AFTER unordered
  ITEM "Review the Q1 financials"
  ITEM "Schedule follow-up meeting"
  ITEM "Update the project timeline"

AT body/paragraph[5]
INSERT LIST AFTER ordered
  ITEM "First, gather requirements"
  ITEM "Second, create the design"
  ITEM "Third, implement and test"
```

| Modifier | Description |
|----------|-------------|
| `ordered` | Numbered list (1, 2, 3...) |
| `unordered` | Bulleted list (default) |

Each `ITEM` line specifies one list item as a quoted string or content
block. Items are added in order.

Nested lists use indented ITEM lines with a sub-list modifier:

```
AT body/heading[text="Overview"]
INSERT LIST AFTER unordered
  ITEM "Backend services"
    ITEM "Authentication" nested
    ITEM "Database" nested
  ITEM "Frontend components"
    ITEM "Dashboard" nested
    ITEM "Settings page" nested
```

Items marked `nested` become children of the preceding non-nested item,
creating a sub-list.

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

#### 5.2.3. SET RUNS

Replaces the content of the addressed element with a sequence of
individually formatted text runs. This enables mixed formatting within
a single paragraph — a requirement for content generated from Markdown
or other rich-text sources.

```
AT body/paragraph[3]
SET RUNS
  RUN "This is "
  RUN "bold text" bold=true
  RUN " and "
  RUN "italic text" italic=true
  RUN " in one paragraph."
```

Each `RUN` line specifies text content as a quoted string, followed by
optional `key=value` formatting properties. Supported properties are the
same as the text (run) properties defined in §7.1.

RUN lines MUST be indented by at least two spaces relative to the SET RUNS
line.

SET RUNS replaces all existing content in the addressed element. After
execution, the element contains exactly the specified runs in order, each
with its declared formatting.

Run-level hyperlinks are specified with the `href` property:

```
AT body/paragraph[1]
SET RUNS
  RUN "Visit "
  RUN "our website" href="https://example.com", color=#2B579A, underline=single
  RUN " for more information."
```

Run-level inline code is expressed with font properties:

```
AT body/paragraph[2]
SET RUNS
  RUN "Use the "
  RUN "officetalk" font-name="Consolas", font-size=10pt, highlight=#F0F0F0
  RUN " command to apply changes."
```

SET RUNS with content blocks for runs containing special characters:

```
AT body/paragraph[4]
SET RUNS
  RUN <<<
Text with "quotes" and other special characters.
>>>
  RUN " followed by " bold=true
  RUN "more text."
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

### 5.5. Annotation Operations

#### 5.5.1. COMMENT

Adds a comment anchored to the addressed element.

```
AT heading[text="Key Risks"]/item[2]
COMMENT "Should we update this figure for Q2?"

AT body/paragraph[text*="revenue"]/run[text*="29%"]
COMMENT "Finance team to verify this number."
```

The comment is attached to the full range of the addressed element. When
the address resolves to a paragraph, the comment spans the entire paragraph.
When the address resolves to a run, the comment spans only that run's text.

COMMENT with content blocks for longer review notes:

```
AT heading[text="Recommendations"]/paragraph[1]
COMMENT <<<
This recommendation conflicts with the budget constraints
outlined in Section 3. Please reconcile before publishing.
>>>
```

The `author` of the comment is implementation-defined. Implementations
SHOULD allow the author to be configured externally (e.g., via a command-line
option or environment variable). If no author is configured, implementations
MAY use a default such as "OfficeTalk".

COMMENT is additive: applying a COMMENT operation to an element that already
has comments MUST NOT remove existing comments. Multiple COMMENT operations
on the same element create multiple distinct comments.

### 5.6. Inspect Operations

#### 5.6.1. INSPECT

Resolves an address against the target document and returns a structured
description of the matched elements.

```
INSPECT <address>
  [DEPTH <integer>]
  [INCLUDE <layer> *("," <layer>)]
  [CONTEXT <integer>]
```

**Modifiers:**

| Keyword   | Type    | Default | Description |
|-----------|---------|---------|-------------|
| `DEPTH`   | integer | 0       | Levels of child elements to include. 0 = matched element only. |
| `INCLUDE` | layers  | (none)  | Comma-separated list: `content`, `properties`, or both. |
| `CONTEXT` | integer | 0       | Number of sibling elements before and after to include. |

When no `INCLUDE` is specified, only addressing information is returned
(element type, position, identity). This is the most lightweight mode —
suitable for confirming an address resolves correctly or mapping document
structure.

#### 5.6.2. Examples

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

#### 5.6.3. Semantics

An OfficeTalk document containing INSPECT operations is a **read document**.
It MUST NOT contain AT blocks, PROPERTY statements, or any write operations.
Implementations MUST reject documents that mix INSPECT and write operations.

Multiple INSPECT operations in a single document are processed sequentially.
Each INSPECT produces one result object in the JSONL response
([Section 14](#14-response-format)).

INSPECT operations do not modify the target document.

#### 5.6.4. Detail Layers

The `INCLUDE` keyword controls which information layers are returned for each
matched element. Layers are additive.

**Addressing (always included):**

Structural identity of the element — type, position within its container,
and type-specific identifiers (sheet name, cell reference, placeholder type,
heading level, style name).

**Content (`INCLUDE content`):**

The textual content of the element. For cells, the display value. For
paragraphs, the full text. For shapes, the text content. For rows, the
cell values. Content is not included by default to keep responses lightweight
when only structure is needed.

**Properties (`INCLUDE properties`):**

Formatting and metadata properties of the element. Font name, size, bold,
italic, color, fill, borders, number format, alignment. Properties are
expensive to compute and verbose in output — only requested when the caller
needs to reason about or modify formatting.

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
| `href` | string | Hyperlink URL (creates clickable link on the run). Use `none` to remove. |

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
differently depending on whether the document is a read document or a write
document.

**Write documents** are processed in three phases: resolution, validation,
and execution.

**Read documents** (containing INSPECT operations) are processed by
resolving each INSPECT address, gathering the requested detail layers, and
producing JSONL response output. INSPECT processing does not modify the target
document.

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

### 8.6. INSPECT Processing

For read documents, the implementation processes each INSPECT block
sequentially:

1. **Resolve** the address against the target document. Unlike write
   operations, INSPECT addresses MAY match multiple elements without
   raising an ambiguity error.
2. **Gather detail layers** — For each matched element, collect the
   addressing layer (always), content layer (if `INCLUDE content` is
   specified), and properties layer (if `INCLUDE properties` is specified).
3. **Expand children** — If `DEPTH > 0`, recursively gather detail layers
   for child elements up to the specified depth.
4. **Gather context** — If `CONTEXT > 0`, include the specified number of
   sibling elements before and after each matched element. Context elements
   include the same detail layers as the matched element.
5. **Include comments** — If the matched element (or its children/context
   elements) has comments, include them in the response regardless of the
   `INCLUDE` setting.
6. **Emit response** — Produce a JSONL response object for the INSPECT
   block (see [Section 14](#14-response-format)).

INSPECT processing MUST NOT modify the target document.

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
- Mixing INSPECT operations with write operations (AT blocks or PROPERTY)
- Invalid INSPECT modifier values (e.g., non-integer DEPTH)
- Unknown INCLUDE layer names (only `content` and `properties` are valid)

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
| `MIXED_OPERATIONS` | Document mixes INSPECT and write operations |

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

### 11.2. Response Media Type Registration

```
Type name:               application
Subtype name:            officetalk-response
Required parameters:     none
Optional parameters:     none
Encoding considerations: UTF-8
Security considerations: See Section 10
Published specification: This document
Applications:            OfficeTalk INSPECT and operation responses
File extension:          .jsonl
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

### 12.8. Word: Document Review with Comments

```
OFFICETALK/1.0
DOCTYPE word

# Flag an outdated statistic for the author to verify
AT heading[text="Financial Summary"]/paragraph[1]/run[text*="29%"]
COMMENT "Please verify this figure — it may have been updated to 31% in the latest filing."

# Request clarification on a vague recommendation
AT heading[text="Recommendations"]/item[2]
COMMENT <<<
This recommendation to "expand by 20%" needs more context:
- What is the baseline headcount?
- What is the budget impact?
- Has HR approved the hiring plan?
>>>

# Note a compliance concern
AT heading[text="Key Risks"]/paragraph[1]
COMMENT "Legal team should review this section before publication."
```

### 12.9. Word: Rich Content with Formatted Runs

```
OFFICETALK/1.0
DOCTYPE word

# Create a paragraph with mixed formatting
AT body/heading[text="Getting Started"]/paragraph[1]
SET RUNS
  RUN "Install the CLI with "
  RUN "dotnet tool install" font-name="Consolas", font-size=10pt, highlight=#F5F5F5
  RUN " and verify with "
  RUN "markmyword version" font-name="Consolas", font-size=10pt, highlight=#F5F5F5
  RUN "."

# Insert a paragraph with a hyperlink
AT body/paragraph[5]
INSERT AFTER "For details, see the specification."
LINK "https://github.com/spec-works/OfficeTalk"
```

### 12.10. Word: Insert Images and Tables

```
OFFICETALK/1.0
DOCTYPE word

# Insert a diagram after the architecture heading
AT body/heading[text="Architecture"]
INSERT IMAGE AFTER "diagrams/system-overview.png"
  alt="System architecture overview"
  width=6in

# Create a comparison table
AT body/heading[text="Comparison"]
INSERT TABLE AFTER rows=4, columns=3
SET CELLS "Feature", "Option A", "Option B"

AT body/table[2]/row[2]
SET CELLS "Performance", "High", "Medium"

AT body/table[2]/row[3]
SET CELLS "Cost", "$500/mo", "$200/mo"

AT body/table[2]/row[4]
SET CELLS "Support", "24/7", "Business hours"
```

### 12.11. Word: Insert Lists

```
OFFICETALK/1.0
DOCTYPE word

# Insert an action items list
AT body/heading[text="Next Steps"]
INSERT LIST AFTER unordered
  ITEM "Review the Q1 financials"
  ITEM "Schedule follow-up meeting"
  ITEM "Update the project timeline"

# Insert a numbered procedure
AT body/heading[text="Installation"]
INSERT LIST AFTER ordered
  ITEM "Download the installer from the releases page"
  ITEM "Run the setup wizard"
  ITEM "Verify the installation"
```

### 12.12. Excel: Inspect Sheet Structure

```
OFFICETALK/1.0
DOCTYPE excel

# Discover what sheets exist
INSPECT sheet[1]

# See all rows with their values
INSPECT sheet["Q1 Budget"]
  DEPTH 1
  INCLUDE content
```

### 12.13. Word: Inspect Document Outline

```
OFFICETALK/1.0
DOCTYPE word

# Get all headings with their text
INSPECT body/heading
  INCLUDE content

# See a heading with surrounding paragraphs for context
INSPECT body/heading[text="Conclusion"]
  INCLUDE content
  CONTEXT 3
```

### 12.14. PowerPoint: Inspect Slide Deck

```
OFFICETALK/1.0
DOCTYPE powerpoint

# Get all slide titles
INSPECT slide
  DEPTH 1
  INCLUDE content

# Full detail on a specific slide
INSPECT slide[3]
  DEPTH 1
  INCLUDE content, properties
```

### 12.15. Excel: Inspect Cell with Properties

```
OFFICETALK/1.0
DOCTYPE excel

# Get cell value and formatting
INSPECT sheet["Q1 Budget"]/D2
  INCLUDE content, properties

# See a cell in context with its neighbors
INSPECT sheet["Q1 Budget"]/D2
  INCLUDE content
  CONTEXT 2
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

block            = ( operation-block / property-block / sheet-block
                   / inspect-block )
operation-block  = address-line 1*( operation-line / comment-line )
address-line     = "AT" [ SP "EACH" ] SP address LF
property-block   = "PROPERTY" SP key "=" value LF
sheet-block      = ("ADD SHEET" / "DELETE SHEET" / "RENAME SHEET") SP quoted-string LF

operation-line   = ( set-op / replace-op / insert-op / delete-op
                   / append-op / prepend-op / format-op / style-op
                   / set-runs-op / link-op
                   / structural-op / comment-op ) LF

; --- Content Operations ---
set-op           = "SET" SP ( quoted-string / content-block / cells-clause )
cells-clause     = "CELLS" SP quoted-string *( "," SP quoted-string )
replace-op       = ["REPLACE" / "REPLACE ALL"] SP quoted-string SP "WITH" SP
                   ( quoted-string / content-block )
insert-op        = ( "INSERT BEFORE" / "INSERT AFTER" ) SP
                   ( quoted-string / content-block / slide-props
                   / image-clause / table-clause / list-clause )
delete-op        = "DELETE" [ SP ( "ROW" / "COLUMN" / "SLIDE" / "SHEET" ) ]
append-op        = "APPEND" SP ( quoted-string / content-block )
prepend-op       = "PREPEND" SP ( quoted-string / content-block )

image-clause     = "IMAGE" SP ( "BEFORE" / "AFTER" ) SP quoted-string
                   *( LF indent image-prop )
image-prop       = ( "alt" / "width" / "height" / "position" ) "=" prop-value

table-clause     = "TABLE" SP ( "BEFORE" / "AFTER" ) SP table-dims
                   *( LF indent table-prop )
table-dims       = "rows=" 1*DIGIT "," SP "columns=" 1*DIGIT
table-prop       = ( "caption" / "width" ) "=" prop-value

list-clause      = "LIST" SP ( "BEFORE" / "AFTER" ) SP [ list-type ]
                   1*( LF indent item-line )
list-type        = "ordered" / "unordered"
item-line        = "ITEM" SP ( quoted-string / content-block ) [ SP "nested" ]

link-op          = "LINK" SP quoted-string

; --- Formatting Operations ---
format-op        = "FORMAT" SP property-list
property-list    = property *( "," SP property )
property         = key "=" prop-value
prop-value       = quoted-string / number / boolean / color / length
style-op         = "STYLE" SP quoted-string

; --- Set Runs Operation ---
set-runs-op      = "SET RUNS" LF 1*( indent run-line LF )
run-line         = "RUN" SP ( quoted-string / content-block )
                   *( SP run-prop )
run-prop         = key "=" prop-value

; --- Structural Operations ---
structural-op    = row-op / column-op / merge-op / slide-op
row-op           = "INSERT ROW" SP ( "BEFORE" / "AFTER" )
column-op        = "INSERT COLUMN" SP ( "BEFORE" / "AFTER" )
merge-op         = "MERGE CELLS TO" SP address
slide-op         = ( "INSERT SLIDE" SP ( "BEFORE" / "AFTER" ) )
                 / "DUPLICATE SLIDE"
slide-props      = *( set-op LF )

; --- Annotation Operations ---
comment-op       = "COMMENT" SP ( quoted-string / content-block )

; --- Inspect Operations ---
inspect-block    = "INSPECT" SP address LF
                   *( inspect-modifier / comment-line )

inspect-modifier = depth-modifier / include-modifier / context-modifier

depth-modifier   = indent "DEPTH" SP 1*DIGIT LF
include-modifier = indent "INCLUDE" SP layer *( "," SP layer ) LF
context-modifier = indent "CONTEXT" SP 1*DIGIT LF

layer            = "content" / "properties"

indent           = 2*WSP

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

## 14. Response Format

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
{"op":"inspect","address":"sheet[\"Q1 Budget\"]","matched":1,"elements":[{"type":"sheet","name":"Q1 Budget","children":[{"type":"excel-row","index":1,"content":{"cells":["Department","Q1 Budget","Q1 Actual","Variance"]}},{"type":"excel-row","index":2,"content":{"cells":["Engineering","150000","162000","-12000"]}}]}]}
```

#### Content and properties — Excel cell

Request:
```
OFFICETALK/1.0
DOCTYPE excel

INSPECT sheet["Q1 Budget"]/D2
  INCLUDE content, properties
```

Response:
```jsonl
{"op":"inspect","address":"sheet[\"Q1 Budget\"]/D2","matched":1,"elements":[{"type":"excel-cell","reference":"D2","sheet":"Q1 Budget","index":4,"of":4,"content":{"value":"-12000","dataType":"string"},"properties":{"bold":false,"font-name":"Calibri","font-size":"11pt","number-format":"General"}}]}
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

#### Word document outline

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
{"op":"inspect","address":"slide","matched":4,"elements":[{"type":"slide","index":1,"of":4,"children":[{"type":"shape","placeholder":"title","content":{"text":"Q1 Business Review"}},{"type":"shape","placeholder":"subtitle","content":{"text":"Prepared by: Finance Team"}}]},{"type":"slide","index":2,"of":4,"children":[{"type":"shape","placeholder":"title","content":{"text":"Revenue Highlights"}},{"type":"shape","placeholder":"body","content":{"text":"Total revenue: $2.4M (+12% YoY)"}}]}]}
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
application/officetalk-response
```

Implementations that accept OfficeTalk documents and produce responses
SHOULD use this media type in Content-Type headers.

### 14.7. Conformance

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
- [RFC 7464] Williams, N., "JavaScript Object Notation (JSON) Text
  Sequences", RFC 7464, February 2015.
- [RFC 8610] Birkholz, H., Vigano, C., and C. Bormann, "Concise Data
  Definition Language (CDDL): A Notational Convention to Express Concise
  Binary Object Representation (CBOR) and JSON Data Structures",
  RFC 8610, June 2019.
- [RFC 9485] Bormann, C. and T. Bray, "I-Regexp: An Interoperable
  Regular Expression Format", RFC 9485, October 2023.
- [ECMA-376] ECMA International, "Office Open XML File Formats", ECMA-376,
  5th Edition, December 2021.
- [Open XML SDK] Microsoft, "Open XML SDK Documentation",
  https://docs.microsoft.com/en-us/office/open-xml/open-xml-sdk
