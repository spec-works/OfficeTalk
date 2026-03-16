using OfficeTalk.Ast;

namespace OfficeTalk.Parsing;

/// <summary>
/// Parses a token stream into an OfficeTalk AST.
/// Supports error recovery — collects errors and continues parsing when possible.
/// </summary>
public class OfficeTalkParser
{
    private readonly List<Token> _tokens;
    private int _pos;
    private readonly List<ParseError> _errors = new();

    public OfficeTalkParser(List<Token> tokens)
    {
        _tokens = tokens ?? throw new ArgumentNullException(nameof(tokens));
        _pos = 0;
    }

    /// <summary>
    /// Parse the token stream into an OfficeTalkDocument AST.
    /// </summary>
    public OfficeTalkDocument Parse()
    {
        var doc = new OfficeTalkDocument();

        // Parse version header
        ParseVersion(doc);
        SkipNewLines();

        // Parse DOCTYPE
        ParseDocType(doc);
        SkipNewLines();

        // Parse blocks
        while (!IsAtEnd())
        {
            SkipNewLines();
            if (IsAtEnd()) break;

            var token = Current();

            if (token.Type == TokenType.Comment)
            {
                Advance(); // skip comments
                continue;
            }

            if (token.Type == TokenType.AT)
            {
                var block = ParseOperationBlock();
                if (block != null)
                    doc.OperationBlocks.Add(block);
            }
            else if (token.Type == TokenType.INSPECT)
            {
                var block = ParseInspectBlock();
                if (block != null)
                    doc.InspectBlocks.Add(block);
            }
            else if (token.Type == TokenType.PROPERTY)
            {
                var prop = ParsePropertySetting();
                if (prop != null)
                    doc.PropertySettings.Add(prop);
            }
            else
            {
                AddError($"Unexpected token '{token.Value}', expected AT, INSPECT, or PROPERTY", token);
                Advance();
            }
        }

        doc.Errors = _errors;
        return doc;
    }

    private void ParseVersion(OfficeTalkDocument doc)
    {
        if (IsAtEnd())
        {
            _errors.Add(new ParseError("Expected OFFICETALK version header", 1, 1));
            return;
        }

        var token = Current();
        if (token.Type != TokenType.Version)
        {
            _errors.Add(new ParseError("Expected OFFICETALK version header", token.Line, token.Column));
            return;
        }

        var value = token.Value;
        if (value.StartsWith("OFFICETALK/", StringComparison.Ordinal))
        {
            doc.Version = value["OFFICETALK/".Length..];
        }
        else
        {
            _errors.Add(new ParseError($"Invalid version header: '{value}'", token.Line, token.Column));
        }

        Advance();
    }

    private void ParseDocType(OfficeTalkDocument doc)
    {
        if (IsAtEnd() || Current().Type != TokenType.DocTypeKeyword)
        {
            _errors.Add(new ParseError("Expected DOCTYPE declaration", CurrentLine(), 1));
            return;
        }

        Advance(); // skip DOCTYPE

        if (IsAtEnd())
        {
            _errors.Add(new ParseError("Expected document type after DOCTYPE", CurrentLine(), 1));
            return;
        }

        var token = Current();
        doc.DocType = token.Type switch
        {
            TokenType.Word => DocType.Word,
            TokenType.Excel => DocType.Excel,
            TokenType.PowerPoint => DocType.PowerPoint,
            _ => HandleInvalidDocType(token)
        };

        Advance();
    }

    private DocType HandleInvalidDocType(Token token)
    {
        _errors.Add(new ParseError($"Invalid document type: '{token.Value}', expected word, excel, or powerpoint", token.Line, token.Column));
        return DocType.Word; // default fallback
    }

    private OperationBlock? ParseOperationBlock()
    {
        var atToken = Current();
        Advance(); // skip AT

        bool isEach = false;
        if (!IsAtEnd() && Current().Type == TokenType.EACH)
        {
            isEach = true;
            Advance();
        }

        var address = ParseAddress();
        if (address == null)
            return null;

        var block = new OperationBlock
        {
            Address = address,
            IsEach = isEach,
            Line = atToken.Line
        };

        SkipNewLines();

        // Parse operations until we hit another AT, PROPERTY, or EOF
        while (!IsAtEnd())
        {
            SkipCommentsAndNewLines();
            if (IsAtEnd()) break;

            var token = Current();
            if (token.Type == TokenType.AT || token.Type == TokenType.PROPERTY)
                break;

            // Check for blank line (two consecutive newlines) which separates blocks
            if (token.Type == TokenType.NewLine)
            {
                Advance();
                continue;
            }

            var op = ParseOperation();
            if (op != null)
                block.Operations.Add(op);
        }

        return block;
    }

    private InspectBlock? ParseInspectBlock()
    {
        var inspectToken = Current();
        Advance(); // skip INSPECT

        var address = ParseAddress();
        if (address == null)
            return null;

        var block = new InspectBlock
        {
            Address = address,
            Line = inspectToken.Line
        };

        SkipNewLines();

        // Parse modifiers (DEPTH, INCLUDE, CONTEXT) on subsequent indented lines
        while (!IsAtEnd())
        {
            SkipCommentsAndNewLines();
            if (IsAtEnd()) break;

            var token = Current();

            if (token.Type == TokenType.DEPTH)
            {
                Advance();
                if (!IsAtEnd() && Current().Type == TokenType.Number)
                {
                    if (int.TryParse(Current().Value, out int depth))
                        block.Depth = depth;
                    else
                        AddError($"Invalid DEPTH value: '{Current().Value}'", Current());
                    Advance();
                }
                else
                {
                    AddError("Expected integer after DEPTH", token);
                }
            }
            else if (token.Type == TokenType.INCLUDE)
            {
                Advance();
                ParseIncludeLayers(block, token);
            }
            else if (token.Type == TokenType.CONTEXT)
            {
                Advance();
                if (!IsAtEnd() && Current().Type == TokenType.Number)
                {
                    if (int.TryParse(Current().Value, out int context))
                        block.Context = context;
                    else
                        AddError($"Invalid CONTEXT value: '{Current().Value}'", Current());
                    Advance();
                }
                else
                {
                    AddError("Expected integer after CONTEXT", token);
                }
            }
            else
            {
                break; // Not a modifier — end of this inspect block
            }
        }

        return block;
    }

    private void ParseIncludeLayers(InspectBlock block, Token includeToken)
    {
        while (!IsAtEnd() && Current().Type != TokenType.NewLine && Current().Type != TokenType.EOF)
        {
            if (Current().Type == TokenType.Comma)
            {
                Advance();
                continue;
            }

            if (Current().Type == TokenType.Identifier || IsSegmentKeyword(Current().Type))
            {
                var layerName = Current().Value.ToLowerInvariant();
                if (layerName == "content")
                {
                    if (!block.Include.Contains(IncludeLayer.Content))
                        block.Include.Add(IncludeLayer.Content);
                    Advance();
                }
                else if (layerName == "properties")
                {
                    if (!block.Include.Contains(IncludeLayer.Properties))
                        block.Include.Add(IncludeLayer.Properties);
                    Advance();
                }
                else
                {
                    AddError($"Unknown INCLUDE layer: '{Current().Value}', expected 'content' or 'properties'", Current());
                    Advance();
                }
            }
            else
            {
                break;
            }
        }

        if (block.Include.Count == 0)
        {
            AddError("Expected at least one layer (content, properties) after INCLUDE", includeToken);
        }
    }

    private PropertySetting? ParsePropertySetting()
    {
        var propToken = Current();
        Advance(); // skip PROPERTY

        if (IsAtEnd() || Current().Type != TokenType.Identifier)
        {
            AddError("Expected property name after PROPERTY", propToken);
            return null;
        }

        var name = Current().Value;
        Advance();

        if (IsAtEnd() || Current().Type != TokenType.Equals)
        {
            AddError("Expected '=' after property name", propToken);
            return null;
        }
        Advance(); // skip =

        if (IsAtEnd() || Current().Type != TokenType.String)
        {
            AddError("Expected string value for property", propToken);
            return null;
        }

        var value = Current().Value;
        Advance();

        return new PropertySetting { Name = name, Value = value, Line = propToken.Line };
    }

    private Address? ParseAddress()
    {
        var address = new Address();
        var rawParts = new List<string>();

        while (!IsAtEnd() && Current().Type != TokenType.NewLine && Current().Type != TokenType.EOF)
        {
            var segment = ParseAddressSegment();
            if (segment == null) break;

            address.Segments.Add(segment);
            rawParts.Add(segment.ToString());

            if (!IsAtEnd() && Current().Type == TokenType.Slash)
            {
                Advance(); // consume /
            }
            else
            {
                break;
            }
        }

        address.RawText = string.Join("/", rawParts);

        if (address.Segments.Count == 0)
        {
            _errors.Add(new ParseError("Expected address after AT", CurrentLine(), 1));
            return null;
        }

        return address;
    }

    private AddressSegment? ParseAddressSegment()
    {
        if (IsAtEnd()) return null;

        var token = Current();

        // The identifier can be a keyword used as segment name (e.g., "body", "row", "slide")
        // or an actual Identifier token
        string identifier;
        if (token.Type == TokenType.Identifier || IsSegmentKeyword(token.Type))
        {
            identifier = token.Value;
            Advance();
        }
        else
        {
            return null;
        }

        var segment = new AddressSegment { Identifier = identifier };

        // Parse predicates in brackets
        if (!IsAtEnd() && Current().Type == TokenType.LeftBracket)
        {
            Advance(); // skip [
            ParsePredicates(segment);

            if (!IsAtEnd() && Current().Type == TokenType.RightBracket)
                Advance(); // skip ]
            else
                _errors.Add(new ParseError("Expected ']' to close predicate", CurrentLine(), CurrentColumn()));
        }

        return segment;
    }

    private void ParsePredicates(AddressSegment segment)
    {
        while (!IsAtEnd() && Current().Type != TokenType.RightBracket)
        {
            var pred = ParsePredicate();
            if (pred != null)
                segment.Predicates.Add(pred);

            if (!IsAtEnd() && Current().Type == TokenType.Comma)
                Advance(); // skip comma separator
        }
    }

    private Predicate? ParsePredicate()
    {
        if (IsAtEnd()) return null;

        var token = Current();

        // Positional: a bare number
        if (token.Type == TokenType.Number)
        {
            Advance();
            if (int.TryParse(token.Value, out int pos))
                return new PositionalPredicate(pos);
            return null;
        }

        // Bare string: "Revenue"
        if (token.Type == TokenType.String)
        {
            Advance();
            return new BareStringPredicate(token.Value);
        }

        // Key-value: key=value, key~=value, etc.
        if (token.Type == TokenType.Identifier || IsSegmentKeyword(token.Type))
        {
            string key = token.Value;
            Advance();

            if (IsAtEnd()) return null;

            var opToken = Current();
            PredicateOperator op;
            switch (opToken.Type)
            {
                case TokenType.Equals:
                    op = PredicateOperator.Equals;
                    break;
                case TokenType.TildeEquals:
                    op = PredicateOperator.TildeEquals;
                    break;
                case TokenType.CaretEquals:
                    op = PredicateOperator.CaretEquals;
                    break;
                case TokenType.DollarEquals:
                    op = PredicateOperator.DollarEquals;
                    break;
                case TokenType.AsteriskEquals:
                    op = PredicateOperator.AsteriskEquals;
                    break;
                default:
                    // Could be a cell ref or bare identifier
                    return null;
            }

            Advance(); // skip operator

            if (IsAtEnd()) return null;

            var valueToken = Current();
            string value;
            if (valueToken.Type == TokenType.String)
            {
                value = valueToken.Value;
                Advance();
            }
            else if (valueToken.Type == TokenType.Number)
            {
                value = valueToken.Value;
                Advance();
            }
            else if (valueToken.Type == TokenType.Boolean)
            {
                value = valueToken.Value;
                Advance();
            }
            else if (valueToken.Type == TokenType.Identifier)
            {
                value = valueToken.Value;
                Advance();
            }
            else
            {
                AddError($"Expected value after operator", valueToken);
                return null;
            }

            return new KeyValuePredicate(key, op, value);
        }

        // Unknown predicate — skip
        Advance();
        return null;
    }

    private Operation? ParseOperation()
    {
        if (IsAtEnd()) return null;

        var token = Current();

        return token.Type switch
        {
            TokenType.SET => ParseSetOrSetCells(),
            TokenType.REPLACE => ParseReplace(),
            TokenType.INSERT => ParseInsert(),
            TokenType.DELETE => ParseDelete(),
            TokenType.APPEND => ParseAppend(),
            TokenType.PREPEND => ParsePrepend(),
            TokenType.FORMAT => ParseFormat(),
            TokenType.STYLE => ParseStyle(),
            TokenType.MERGE => ParseMerge(),
            TokenType.DUPLICATE => ParseDuplicate(),
            TokenType.RENAME => ParseRename(),
            TokenType.ADD => ParseAdd(),
            TokenType.COMMENT_KW => ParseComment(),
            _ => HandleUnexpectedOperation(token)
        };
    }

    private Operation? ParseSetOrSetCells()
    {
        var setToken = Current();
        Advance(); // skip SET

        if (!IsAtEnd() && Current().Type == TokenType.CELLS)
        {
            return ParseSetCellsAfterSet(setToken);
        }

        var content = ParseContentValue();
        if (content == null)
        {
            AddError("Expected content after SET", setToken);
            return null;
        }

        return new SetOperation { Content = content, Line = setToken.Line };
    }

    private Operation? ParseSetCellsAfterSet(Token setToken)
    {
        Advance(); // skip CELLS

        var values = new List<string>();
        while (!IsAtEnd() && Current().Type == TokenType.String)
        {
            values.Add(Current().Value);
            Advance();

            if (!IsAtEnd() && Current().Type == TokenType.Comma)
                Advance();
        }

        return new SetCellsOperation { Values = values, Line = setToken.Line };
    }

    private Operation? ParseReplace()
    {
        var replToken = Current();
        Advance(); // skip REPLACE

        bool isAll = false;
        if (!IsAtEnd() && Current().Type == TokenType.ALL)
        {
            isAll = true;
            Advance();
        }

        if (IsAtEnd() || Current().Type != TokenType.String)
        {
            AddError("Expected search string after REPLACE", replToken);
            return null;
        }

        var search = Current().Value;
        Advance();

        if (IsAtEnd() || Current().Type != TokenType.WITH)
        {
            AddError("Expected WITH after search string", replToken);
            return null;
        }
        Advance(); // skip WITH

        if (IsAtEnd() || Current().Type != TokenType.String)
        {
            AddError("Expected replacement string after WITH", replToken);
            return null;
        }

        var replacement = Current().Value;
        Advance();

        return new ReplaceOperation
        {
            Search = search,
            Replacement = replacement,
            IsAll = isAll,
            Line = replToken.Line
        };
    }

    private Operation? ParseInsert()
    {
        var insToken = Current();
        Advance(); // skip INSERT

        if (IsAtEnd())
        {
            AddError("Expected BEFORE, AFTER, ROW, COLUMN, or SLIDE after INSERT", insToken);
            return null;
        }

        var next = Current();

        // INSERT ROW BEFORE/AFTER
        if (next.Type == TokenType.ROW)
        {
            Advance();
            var pos = ParseInsertPosition(insToken);
            return pos.HasValue ? new InsertRowOperation { Position = pos.Value, Line = insToken.Line } : null;
        }

        // INSERT COLUMN BEFORE/AFTER
        if (next.Type == TokenType.COLUMN)
        {
            Advance();
            var pos = ParseInsertPosition(insToken);
            return pos.HasValue ? new InsertColumnOperation { Position = pos.Value, Line = insToken.Line } : null;
        }

        // INSERT SLIDE BEFORE/AFTER
        if (next.Type == TokenType.SLIDE)
        {
            Advance();
            var pos = ParseInsertPosition(insToken);
            return pos.HasValue ? new InsertSlideOperation { Position = pos.Value, Line = insToken.Line } : null;
        }

        // INSERT BEFORE content
        if (next.Type == TokenType.BEFORE)
        {
            Advance();
            var content = ParseContentValue();
            if (content == null)
            {
                AddError("Expected content after INSERT BEFORE", insToken);
                return null;
            }
            return new InsertBeforeOperation { Content = content, Line = insToken.Line };
        }

        // INSERT AFTER content
        if (next.Type == TokenType.AFTER)
        {
            Advance();
            var content = ParseContentValue();
            if (content == null)
            {
                AddError("Expected content after INSERT AFTER", insToken);
                return null;
            }
            return new InsertAfterOperation { Content = content, Line = insToken.Line };
        }

        AddError($"Unexpected token '{next.Value}' after INSERT", next);
        return null;
    }

    private InsertPosition? ParseInsertPosition(Token context)
    {
        if (IsAtEnd())
        {
            AddError("Expected BEFORE or AFTER", context);
            return null;
        }

        var token = Current();
        if (token.Type == TokenType.BEFORE)
        {
            Advance();
            return Ast.InsertPosition.Before;
        }
        if (token.Type == TokenType.AFTER)
        {
            Advance();
            return Ast.InsertPosition.After;
        }

        AddError($"Expected BEFORE or AFTER, got '{token.Value}'", token);
        return null;
    }

    private Operation? ParseDelete()
    {
        var delToken = Current();
        Advance(); // skip DELETE

        var target = DeleteTarget.Element;

        if (!IsAtEnd())
        {
            var next = Current();
            if (next.Type == TokenType.ROW) { target = DeleteTarget.Row; Advance(); }
            else if (next.Type == TokenType.COLUMN) { target = DeleteTarget.Column; Advance(); }
            else if (next.Type == TokenType.SLIDE) { target = DeleteTarget.Slide; Advance(); }
            else if (next.Type == TokenType.SHEET) { target = DeleteTarget.Sheet; Advance(); }
        }

        return new DeleteOperation { Target = target, Line = delToken.Line };
    }

    private Operation? ParseAppend()
    {
        var token = Current();
        Advance(); // skip APPEND

        var content = ParseContentValue();
        if (content == null)
        {
            AddError("Expected content after APPEND", token);
            return null;
        }

        return new AppendOperation { Content = content, Line = token.Line };
    }

    private Operation? ParsePrepend()
    {
        var token = Current();
        Advance(); // skip PREPEND

        var content = ParseContentValue();
        if (content == null)
        {
            AddError("Expected content after PREPEND", token);
            return null;
        }

        return new PrependOperation { Content = content, Line = token.Line };
    }

    private Operation? ParseFormat()
    {
        var fmtToken = Current();
        Advance(); // skip FORMAT

        var properties = new Dictionary<string, object>();

        while (!IsAtEnd() && Current().Type != TokenType.NewLine && Current().Type != TokenType.EOF
            && Current().Type != TokenType.AT && Current().Type != TokenType.PROPERTY)
        {
            if (Current().Type == TokenType.Comment)
            {
                Advance();
                break;
            }

            if (Current().Type == TokenType.Comma)
            {
                Advance();
                continue;
            }

            // Expect: identifier = value
            if (Current().Type != TokenType.Identifier && !IsSegmentKeyword(Current().Type))
                break;

            var key = Current().Value;
            Advance();

            if (IsAtEnd() || Current().Type != TokenType.Equals)
            {
                AddError($"Expected '=' after format property '{key}'", fmtToken);
                break;
            }
            Advance(); // skip =

            if (IsAtEnd())
            {
                AddError($"Expected value for format property '{key}'", fmtToken);
                break;
            }

            object value = ParseFormatValue();
            properties[key] = value;
        }

        return new FormatOperation { Properties = properties, Line = fmtToken.Line };
    }

    private object ParseFormatValue()
    {
        var token = Current();
        Advance();

        return token.Type switch
        {
            TokenType.String => token.Value,
            TokenType.Number => double.TryParse(token.Value, out var d) ? d : token.Value,
            TokenType.Boolean => token.Value == "true",
            TokenType.Color => token.Value,
            TokenType.Length => token.Value,
            TokenType.Identifier => token.Value,
            _ => token.Value
        };
    }

    private Operation? ParseStyle()
    {
        var styleToken = Current();
        Advance(); // skip STYLE

        if (IsAtEnd() || Current().Type != TokenType.String)
        {
            AddError("Expected style name string after STYLE", styleToken);
            return null;
        }

        var name = Current().Value;
        Advance();

        return new StyleOperation { StyleName = name, Line = styleToken.Line };
    }

    private Operation? ParseMerge()
    {
        var mergeToken = Current();
        Advance(); // skip MERGE

        if (IsAtEnd() || Current().Type != TokenType.CELLS)
        {
            AddError("Expected CELLS after MERGE", mergeToken);
            return null;
        }
        Advance(); // skip CELLS

        if (IsAtEnd() || Current().Type != TokenType.TO)
        {
            AddError("Expected TO after MERGE CELLS", mergeToken);
            return null;
        }
        Advance(); // skip TO

        var address = ParseAddress();
        if (address == null)
        {
            AddError("Expected address after MERGE CELLS TO", mergeToken);
            return null;
        }

        return new MergeCellsOperation { TargetAddress = address, Line = mergeToken.Line };
    }

    private Operation? ParseDuplicate()
    {
        var dupToken = Current();
        Advance(); // skip DUPLICATE

        if (!IsAtEnd() && Current().Type == TokenType.SLIDE)
            Advance(); // skip SLIDE

        return new DuplicateSlideOperation { Line = dupToken.Line };
    }

    private Operation? ParseRename()
    {
        var renToken = Current();
        Advance(); // skip RENAME

        if (!IsAtEnd() && Current().Type == TokenType.SHEET)
            Advance(); // skip SHEET

        if (IsAtEnd() || Current().Type != TokenType.String)
        {
            AddError("Expected new name string after RENAME SHEET", renToken);
            return null;
        }

        var newName = Current().Value;
        Advance();

        return new RenameSheetOperation { NewName = newName, Line = renToken.Line };
    }

    private Operation? ParseAdd()
    {
        var addToken = Current();
        Advance(); // skip ADD

        if (!IsAtEnd() && Current().Type == TokenType.SHEET)
            Advance(); // skip SHEET

        if (IsAtEnd() || Current().Type != TokenType.String)
        {
            AddError("Expected sheet name string after ADD SHEET", addToken);
            return null;
        }

        var name = Current().Value;
        Advance();

        return new AddSheetOperation { Name = name, Line = addToken.Line };
    }

    private Operation? ParseComment()
    {
        var token = Current();
        Advance(); // skip COMMENT

        var content = ParseContentValue();
        if (content == null)
        {
            AddError("Expected comment text after COMMENT", token);
            return null;
        }

        return new CommentOperation { Content = content, Line = token.Line };
    }

    private ContentValue? ParseContentValue()
    {
        if (IsAtEnd()) return null;

        var token = Current();

        if (token.Type == TokenType.String)
        {
            Advance();
            return new ContentValue(token.Value, false);
        }

        if (token.Type == TokenType.ContentBlock)
        {
            Advance();
            return new ContentValue(token.Value, true);
        }

        return null;
    }

    private Operation? HandleUnexpectedOperation(Token token)
    {
        AddError($"Unexpected token '{token.Value}', expected an operation keyword", token);
        Advance();
        return null;
    }

    private void SkipNewLines()
    {
        while (!IsAtEnd() && Current().Type == TokenType.NewLine)
            Advance();
    }

    private void SkipCommentsAndNewLines()
    {
        while (!IsAtEnd() && (Current().Type == TokenType.NewLine || Current().Type == TokenType.Comment))
            Advance();
    }

    private bool IsAtEnd() => _pos >= _tokens.Count || _tokens[_pos].Type == TokenType.EOF;

    private Token Current() => _pos < _tokens.Count ? _tokens[_pos] : new Token(TokenType.EOF, "", 0, 0);

    private void Advance()
    {
        if (_pos < _tokens.Count)
            _pos++;
    }

    private int CurrentLine() => _pos < _tokens.Count ? _tokens[_pos].Line : 0;
    private int CurrentColumn() => _pos < _tokens.Count ? _tokens[_pos].Column : 0;

    private void AddError(string message, Token token) =>
        _errors.Add(new ParseError(message, token.Line, token.Column));

    private static bool IsSegmentKeyword(TokenType type) =>
        type is TokenType.ROW or TokenType.COLUMN or TokenType.SLIDE or TokenType.SHEET
            or TokenType.STYLE or TokenType.Identifier or TokenType.Word or TokenType.Excel
            or TokenType.PowerPoint or TokenType.CELLS or TokenType.BEFORE or TokenType.AFTER
            or TokenType.DELETE or TokenType.ADD or TokenType.RENAME or TokenType.DUPLICATE
            or TokenType.SET or TokenType.FORMAT or TokenType.REPLACE or TokenType.INSERT
            or TokenType.APPEND or TokenType.PREPEND or TokenType.MERGE or TokenType.TO
            or TokenType.ALL or TokenType.EACH or TokenType.AT or TokenType.WITH
            or TokenType.PROPERTY or TokenType.DocTypeKeyword
            or TokenType.DEPTH or TokenType.INCLUDE or TokenType.CONTEXT;
}
