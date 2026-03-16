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
            TokenType.SET => ParseSetOrSetCellsOrSetRuns(),
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
            TokenType.LINK => ParseLink(),
            _ => HandleUnexpectedOperation(token)
        };
    }

    private Operation? ParseSetOrSetCellsOrSetRuns()
    {
        var setToken = Current();
        Advance(); // skip SET

        if (!IsAtEnd() && Current().Type == TokenType.CELLS)
        {
            return ParseSetCellsAfterSet(setToken);
        }

        if (!IsAtEnd() && Current().Type == TokenType.RUNS)
        {
            return ParseSetRunsAfterSet(setToken);
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
            AddError("Expected BEFORE, AFTER, ROW, COLUMN, SLIDE, IMAGE, TABLE, or LIST after INSERT", insToken);
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

        // INSERT IMAGE BEFORE/AFTER "source"
        if (next.Type == TokenType.IMAGE)
        {
            Advance();
            return ParseInsertImage(insToken);
        }

        // INSERT TABLE BEFORE/AFTER rows=N, columns=M
        if (next.Type == TokenType.TABLE)
        {
            Advance();
            return ParseInsertTable(insToken);
        }

        // INSERT LIST BEFORE/AFTER [ordered|unordered]
        if (next.Type == TokenType.LIST)
        {
            Advance();
            return ParseInsertList(insToken);
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

    private Operation? ParseLink()
    {
        var linkToken = Current();
        Advance(); // skip LINK

        if (IsAtEnd() || Current().Type != TokenType.String)
        {
            AddError("Expected URL string after LINK", linkToken);
            return null;
        }

        var url = Current().Value;
        Advance();

        return new LinkOperation { Url = url, Line = linkToken.Line };
    }

    private Operation? ParseInsertImage(Token insToken)
    {
        var pos = ParseInsertPosition(insToken);
        if (!pos.HasValue) return null;

        if (IsAtEnd() || Current().Type != TokenType.String)
        {
            AddError("Expected image source string after INSERT IMAGE BEFORE/AFTER", insToken);
            return null;
        }

        var source = Current().Value;
        Advance();

        // Parse optional indented properties (alt, width, height, position) on next lines
        var properties = ParseIndentedProperties();

        return new InsertImageOperation
        {
            Position = pos.Value,
            Source = source,
            Properties = properties,
            Line = insToken.Line
        };
    }

    private Operation? ParseInsertTable(Token insToken)
    {
        var pos = ParseInsertPosition(insToken);
        if (!pos.HasValue) return null;

        // Parse rows=N, columns=M as key=value pairs
        int rows = 0, columns = 0;
        var properties = new Dictionary<string, object>();

        // Parse inline key=value pairs: rows=N, columns=M
        while (!IsAtEnd() && Current().Type != TokenType.NewLine && Current().Type != TokenType.EOF)
        {
            if (Current().Type == TokenType.Comma)
            {
                Advance();
                continue;
            }

            if (Current().Type != TokenType.Identifier && !IsSegmentKeyword(Current().Type))
                break;

            var key = Current().Value;
            Advance();

            if (IsAtEnd() || Current().Type != TokenType.Equals)
            {
                AddError($"Expected '=' after table property '{key}'", insToken);
                break;
            }
            Advance(); // skip =

            if (IsAtEnd())
            {
                AddError($"Expected value for table property '{key}'", insToken);
                break;
            }

            var value = ParseFormatValue();

            if (key.Equals("rows", StringComparison.OrdinalIgnoreCase) && value is double d1)
                rows = (int)d1;
            else if (key.Equals("columns", StringComparison.OrdinalIgnoreCase) && value is double d2)
                columns = (int)d2;
            else
                properties[key] = value;
        }

        // Also parse indented properties on next lines
        var indentedProps = ParseIndentedProperties();
        foreach (var kvp in indentedProps)
        {
            if (kvp.Key.Equals("rows", StringComparison.OrdinalIgnoreCase) && kvp.Value is double d3)
                rows = (int)d3;
            else if (kvp.Key.Equals("columns", StringComparison.OrdinalIgnoreCase) && kvp.Value is double d4)
                columns = (int)d4;
            else
                properties[kvp.Key] = kvp.Value;
        }

        if (rows <= 0 || columns <= 0)
        {
            AddError("INSERT TABLE requires rows and columns with positive values", insToken);
            return null;
        }

        return new InsertTableOperation
        {
            Position = pos.Value,
            Rows = rows,
            Columns = columns,
            Properties = properties,
            Line = insToken.Line
        };
    }

    private Operation? ParseInsertList(Token insToken)
    {
        var pos = ParseInsertPosition(insToken);
        if (!pos.HasValue) return null;

        // Parse optional list type (ordered/unordered)
        var listType = Ast.ListType.Unordered;
        if (!IsAtEnd() && Current().Type == TokenType.Identifier)
        {
            var typeValue = Current().Value.ToLowerInvariant();
            if (typeValue == "ordered")
            {
                listType = Ast.ListType.Ordered;
                Advance();
            }
            else if (typeValue == "unordered")
            {
                listType = Ast.ListType.Unordered;
                Advance();
            }
        }

        // Parse ITEM lines (on subsequent lines, indented)
        var items = new List<ListItem>();
        SkipNewLines();

        while (!IsAtEnd() && Current().Type == TokenType.ITEM)
        {
            Advance(); // skip ITEM

            var content = ParseContentValue();
            if (content == null)
            {
                AddError("Expected content after ITEM", Current());
                break;
            }

            bool isNested = false;
            if (!IsAtEnd() && Current().Type == TokenType.NESTED)
            {
                isNested = true;
                Advance();
            }

            items.Add(new ListItem { Content = content, IsNested = isNested });
            SkipNewLines();
        }

        return new InsertListOperation
        {
            Position = pos.Value,
            ListType = listType,
            Items = items,
            Line = insToken.Line
        };
    }

    private Operation? ParseSetRunsAfterSet(Token setToken)
    {
        Advance(); // skip RUNS

        var runs = new List<RunDefinition>();
        SkipNewLines();

        while (!IsAtEnd() && Current().Type == TokenType.RUN)
        {
            Advance(); // skip RUN

            var content = ParseContentValue();
            if (content == null)
            {
                AddError("Expected content after RUN", setToken);
                break;
            }

            // Parse optional inline properties: key=value, key=value
            var properties = new Dictionary<string, object>();
            while (!IsAtEnd() && Current().Type != TokenType.NewLine && Current().Type != TokenType.EOF
                && Current().Type != TokenType.RUN)
            {
                if (Current().Type == TokenType.Comma)
                {
                    Advance();
                    continue;
                }

                if (Current().Type != TokenType.Identifier && !IsSegmentKeyword(Current().Type))
                    break;

                var key = Current().Value;
                Advance();

                if (IsAtEnd() || Current().Type != TokenType.Equals)
                {
                    AddError($"Expected '=' after run property '{key}'", setToken);
                    break;
                }
                Advance(); // skip =

                if (IsAtEnd())
                {
                    AddError($"Expected value for run property '{key}'", setToken);
                    break;
                }

                var value = ParseFormatValue();
                properties[key] = value;
            }

            runs.Add(new RunDefinition { Content = content, Properties = properties });
            SkipNewLines();
        }

        return new SetRunsOperation { Runs = runs, Line = setToken.Line };
    }

    /// <summary>
    /// Parse indented key=value properties on subsequent lines.
    /// Used by INSERT IMAGE and INSERT TABLE for multi-line property blocks.
    /// Stops when it hits a non-property line (another operation keyword, AT, etc.)
    /// </summary>
    private Dictionary<string, object> ParseIndentedProperties()
    {
        var properties = new Dictionary<string, object>();

        // Save position to peek ahead
        while (true)
        {
            // Skip newlines
            int savedPos = _pos;
            SkipNewLines();

            if (IsAtEnd()) break;

            var token = Current();
            // If the next token is an identifier that could be a property key followed by =,
            // parse it. Otherwise, restore position and stop.
            if ((token.Type == TokenType.Identifier || IsSegmentKeyword(token.Type))
                && !IsOperationKeyword(token.Type))
            {
                // Peek ahead for = sign
                int peekPos = _pos + 1;
                if (peekPos < _tokens.Count && _tokens[peekPos].Type == TokenType.Equals)
                {
                    var key = token.Value;
                    Advance(); // skip key
                    Advance(); // skip =

                    if (!IsAtEnd())
                    {
                        var value = ParseFormatValue();
                        properties[key] = value;
                        continue;
                    }
                }
            }

            // Not a property — restore position and stop
            _pos = savedPos;
            break;
        }

        return properties;
    }

    private static bool IsOperationKeyword(TokenType type) =>
        type is TokenType.SET or TokenType.REPLACE or TokenType.INSERT or TokenType.DELETE
            or TokenType.APPEND or TokenType.PREPEND or TokenType.FORMAT or TokenType.STYLE
            or TokenType.MERGE or TokenType.DUPLICATE or TokenType.RENAME or TokenType.ADD
            or TokenType.COMMENT_KW or TokenType.LINK or TokenType.AT or TokenType.INSPECT
            or TokenType.PROPERTY;

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
            or TokenType.DEPTH or TokenType.INCLUDE or TokenType.CONTEXT
            or TokenType.LINK or TokenType.IMAGE or TokenType.TABLE
            or TokenType.LIST or TokenType.ITEM or TokenType.RUNS
            or TokenType.RUN or TokenType.NESTED;
}
