# OfficeTalk Test Cases

Shared, language-independent test cases for the OfficeTalk component.

## Format

Each test case consists of an `.otk` input file containing an OfficeTalk document,
paired with a `.json` file describing the expected parse result (AST structure).

### Positive Tests

Files in the root of this directory are valid OfficeTalk documents that parsers
MUST accept and parse correctly.

| File | Description |
|------|-------------|
| `simple-set.otk` / `.json` | Minimal document with a single SET operation |
| `replace-text.otk` / `.json` | REPLACE operation with WITH clause |
| `replace-all.otk` / `.json` | REPLACE ALL operation for bulk replacement |
| `insert-content-block.otk` / `.json` | INSERT BEFORE with multi-line content block |
| `delete-element.otk` / `.json` | DELETE operation on a paragraph |
| `format-properties.otk` / `.json` | FORMAT with multiple properties |
| `style-operation.otk` / `.json` | STYLE operation applying a named style |
| `append-prepend.otk` / `.json` | APPEND and PREPEND operations |
| `document-properties.otk` / `.json` | PROPERTY lines setting document metadata |
| `multiple-blocks.otk` / `.json` | Document with multiple operation blocks |
| `at-each.otk` / `.json` | AT EACH modifier for bulk operations |
| `text-predicates.otk` / `.json` | Various text matching predicates (exact, contains, starts-with, ends-with, regex) |
| `table-operations.otk` / `.json` | Table structural operations (INSERT ROW, SET CELLS) |
| `comments-and-whitespace.otk` / `.json` | Comments, blank lines, mixed whitespace |
| `all-address-types.otk` / `.json` | Comprehensive address syntax coverage |
| `spec-example-12-1.otk` / `.json` | Example 12.1 from the OfficeTalk specification |
| `spec-example-12-6.otk` / `.json` | Example 12.6 from the OfficeTalk specification (comprehensive rewrite) |

### Negative Tests

Files in the `negative/` subdirectory are invalid OfficeTalk documents that parsers
SHOULD reject with appropriate errors.

| File | Description |
|------|-------------|
| `missing-header.otk` | Missing OFFICETALK version line |
| `missing-doctype.otk` | Missing DOCTYPE line |
| `invalid-doctype.otk` | Invalid DOCTYPE value |
| `missing-at.otk` | Operation without AT address |
| `unterminated-string.otk` | Unterminated quoted string |
| `unterminated-content-block.otk` | Content block missing closing `>>>` |
| `invalid-address.otk` | Malformed address syntax |
| `unknown-operation.otk` | Unrecognized operation keyword |
