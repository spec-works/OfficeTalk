using System.Text;

namespace OfficeTalk.Parsing;

/// <summary>
/// Line-oriented lexer for the OfficeTalk grammar.
/// Tokenizes input text into a sequence of tokens for the parser.
/// </summary>
public class OfficeTalkLexer
{
    private readonly string _source;
    private int _pos;
    private int _line;
    private int _column;
    private readonly List<Token> _tokens = new();

    private static readonly Dictionary<string, TokenType> Keywords = new(StringComparer.OrdinalIgnoreCase)
    {
        ["DOCTYPE"] = TokenType.DocTypeKeyword,
        ["AT"] = TokenType.AT,
        ["EACH"] = TokenType.EACH,
        ["SET"] = TokenType.SET,
        ["REPLACE"] = TokenType.REPLACE,
        ["WITH"] = TokenType.WITH,
        ["ALL"] = TokenType.ALL,
        ["INSERT"] = TokenType.INSERT,
        ["BEFORE"] = TokenType.BEFORE,
        ["AFTER"] = TokenType.AFTER,
        ["DELETE"] = TokenType.DELETE,
        ["APPEND"] = TokenType.APPEND,
        ["PREPEND"] = TokenType.PREPEND,
        ["FORMAT"] = TokenType.FORMAT,
        ["STYLE"] = TokenType.STYLE,
        ["PROPERTY"] = TokenType.PROPERTY,
        ["MERGE"] = TokenType.MERGE,
        ["CELLS"] = TokenType.CELLS,
        ["TO"] = TokenType.TO,
        ["ROW"] = TokenType.ROW,
        ["COLUMN"] = TokenType.COLUMN,
        ["SLIDE"] = TokenType.SLIDE,
        ["SHEET"] = TokenType.SHEET,
        ["ADD"] = TokenType.ADD,
        ["RENAME"] = TokenType.RENAME,
        ["DUPLICATE"] = TokenType.DUPLICATE,
    };

    private static readonly Dictionary<string, TokenType> DocTypeValues = new(StringComparer.OrdinalIgnoreCase)
    {
        ["word"] = TokenType.Word,
        ["excel"] = TokenType.Excel,
        ["powerpoint"] = TokenType.PowerPoint,
    };

    public OfficeTalkLexer(string source)
    {
        _source = source ?? throw new ArgumentNullException(nameof(source));
        _pos = 0;
        _line = 1;
        _column = 1;
    }

    /// <summary>
    /// Tokenize the entire source and return all tokens.
    /// </summary>
    public List<Token> Tokenize()
    {
        _tokens.Clear();
        _pos = 0;
        _line = 1;
        _column = 1;

        // First line must be version header
        TokenizeVersionLine();

        // Tokenize remaining lines
        while (_pos < _source.Length)
        {
            TokenizeNext();
        }

        _tokens.Add(new Token(TokenType.EOF, "", _line, _column));
        return _tokens;
    }

    private void TokenizeVersionLine()
    {
        SkipWhitespaceInLine();

        if (_pos >= _source.Length)
        {
            _tokens.Add(new Token(TokenType.EOF, "", _line, _column));
            return;
        }

        int startLine = _line;
        int startCol = _column;

        // Look for OFFICETALK/x.y
        var lineText = ReadToEndOfLine();
        if (lineText.StartsWith("OFFICETALK/", StringComparison.Ordinal))
        {
            _tokens.Add(new Token(TokenType.Version, lineText, startLine, startCol));
        }
        else
        {
            // Not a valid header — still emit as version token with the content for parser to handle
            _tokens.Add(new Token(TokenType.Version, lineText, startLine, startCol));
        }

        ConsumeNewLine();
    }

    private void TokenizeNext()
    {
        if (_pos >= _source.Length)
            return;

        char c = _source[_pos];

        // Newlines
        if (c == '\n')
        {
            _tokens.Add(new Token(TokenType.NewLine, "\n", _line, _column));
            Advance();
            _line++;
            _column = 1;
            return;
        }

        if (c == '\r')
        {
            Advance();
            if (_pos < _source.Length && _source[_pos] == '\n')
                Advance();
            _tokens.Add(new Token(TokenType.NewLine, "\n", _line, _column));
            _line++;
            _column = 1;
            return;
        }

        // Skip whitespace (not newlines)
        if (c == ' ' || c == '\t')
        {
            SkipWhitespaceInLine();
            return;
        }

        // Colors (# followed by hex digit) — must be checked before comments
        if (c == '#' && _pos + 1 < _source.Length && IsHexDigit(_source[_pos + 1]))
        {
            TokenizeColor();
            return;
        }

        // Comments (# not followed by hex digit)
        if (c == '#')
        {
            int startCol = _column;
            var comment = ReadToEndOfLine();
            _tokens.Add(new Token(TokenType.Comment, comment, _line, startCol));
            return;
        }

        // Content block start <<<
        if (c == '<' && Peek(1) == '<' && Peek(2) == '<')
        {
            TokenizeContentBlock();
            return;
        }

        // Structural characters
        if (c == '/')
        {
            _tokens.Add(new Token(TokenType.Slash, "/", _line, _column));
            Advance();
            return;
        }

        if (c == '[')
        {
            _tokens.Add(new Token(TokenType.LeftBracket, "[", _line, _column));
            Advance();
            return;
        }

        if (c == ']')
        {
            _tokens.Add(new Token(TokenType.RightBracket, "]", _line, _column));
            Advance();
            return;
        }

        if (c == ',')
        {
            _tokens.Add(new Token(TokenType.Comma, ",", _line, _column));
            Advance();
            return;
        }

        // Operators: =, ~=, ^=, $=, *=
        if (c == '~' && Peek(1) == '=')
        {
            _tokens.Add(new Token(TokenType.TildeEquals, "~=", _line, _column));
            Advance(); Advance();
            return;
        }

        if (c == '^' && Peek(1) == '=')
        {
            _tokens.Add(new Token(TokenType.CaretEquals, "^=", _line, _column));
            Advance(); Advance();
            return;
        }

        if (c == '$' && Peek(1) == '=')
        {
            _tokens.Add(new Token(TokenType.DollarEquals, "$=", _line, _column));
            Advance(); Advance();
            return;
        }

        if (c == '*' && Peek(1) == '=')
        {
            _tokens.Add(new Token(TokenType.AsteriskEquals, "*=", _line, _column));
            Advance(); Advance();
            return;
        }

        if (c == '=')
        {
            _tokens.Add(new Token(TokenType.Equals, "=", _line, _column));
            Advance();
            return;
        }

        // Strings
        if (c == '"')
        {
            TokenizeString();
            return;
        }

        // Numbers (and possibly lengths like 12pt, 1.5in)
        if (char.IsDigit(c) || (c == '-' && _pos + 1 < _source.Length && char.IsDigit(_source[_pos + 1])))
        {
            TokenizeNumberOrLength();
            return;
        }

        // Identifiers and keywords
        if (char.IsLetter(c) || c == '_' || c == '-')
        {
            TokenizeIdentifierOrKeyword();
            return;
        }

        // Unknown character — skip
        Advance();
    }

    private void TokenizeContentBlock()
    {
        int startLine = _line;
        int startCol = _column;

        // Consume <<<
        Advance(); Advance(); Advance();

        // Skip rest of <<< line (should be newline)
        SkipToNewLine();
        ConsumeNewLine();

        var sb = new StringBuilder();
        bool foundEnd = false;

        while (_pos < _source.Length)
        {
            // Check if current line starts with >>>
            if (_source[_pos] == '>' && Peek(1) == '>' && Peek(2) == '>')
            {
                // Consume >>>
                Advance(); Advance(); Advance();
                SkipToNewLine();
                foundEnd = true;
                break;
            }

            // Read the line content
            var lineContent = ReadToEndOfLine();
            if (sb.Length > 0)
                sb.Append('\n');
            sb.Append(lineContent);
            ConsumeNewLine();
        }

        // Strip leading/trailing blank lines from content
        var content = sb.ToString().Trim('\n', '\r');

        _tokens.Add(new Token(TokenType.ContentBlock, content, startLine, startCol));

        if (!foundEnd)
        {
            // Unterminated content block — parser will handle the error
        }
    }

    private void TokenizeString()
    {
        int startLine = _line;
        int startCol = _column;
        Advance(); // skip opening "

        var sb = new StringBuilder();
        while (_pos < _source.Length && _source[_pos] != '"')
        {
            if (_source[_pos] == '\\' && _pos + 1 < _source.Length)
            {
                char next = _source[_pos + 1];
                switch (next)
                {
                    case '"': sb.Append('"'); Advance(); Advance(); break;
                    case '\\': sb.Append('\\'); Advance(); Advance(); break;
                    case 'n': sb.Append('\n'); Advance(); Advance(); break;
                    case 't': sb.Append('\t'); Advance(); Advance(); break;
                    default: sb.Append(_source[_pos]); Advance(); break;
                }
            }
            else if (_source[_pos] == '\n' || _source[_pos] == '\r')
            {
                break; // Unterminated string
            }
            else
            {
                sb.Append(_source[_pos]);
                Advance();
            }
        }

        if (_pos < _source.Length && _source[_pos] == '"')
            Advance(); // skip closing "

        _tokens.Add(new Token(TokenType.String, sb.ToString(), startLine, startCol));
    }

    private void TokenizeColor()
    {
        int startLine = _line;
        int startCol = _column;
        Advance(); // skip #

        var sb = new StringBuilder("#");
        while (_pos < _source.Length && IsHexDigit(_source[_pos]))
        {
            sb.Append(_source[_pos]);
            Advance();
        }

        _tokens.Add(new Token(TokenType.Color, sb.ToString(), startLine, startCol));
    }

    private void TokenizeNumberOrLength()
    {
        int startLine = _line;
        int startCol = _column;
        var sb = new StringBuilder();

        if (_source[_pos] == '-')
        {
            sb.Append('-');
            Advance();
        }

        while (_pos < _source.Length && (char.IsDigit(_source[_pos]) || _source[_pos] == '.'))
        {
            sb.Append(_source[_pos]);
            Advance();
        }

        // Check for unit suffix (pt, in, cm, emu, %)
        if (_pos < _source.Length)
        {
            string? unit = null;
            if (_source[_pos] == '%')
            {
                unit = "%";
                sb.Append('%');
                Advance();
            }
            else if (_pos + 1 < _source.Length)
            {
                var twoChar = _source.Substring(_pos, 2);
                if (twoChar is "pt" or "in" or "cm")
                {
                    unit = twoChar;
                    sb.Append(twoChar);
                    Advance(); Advance();
                }
                else if (_pos + 2 < _source.Length && _source.Substring(_pos, 3) == "emu")
                {
                    unit = "emu";
                    sb.Append("emu");
                    Advance(); Advance(); Advance();
                }
            }

            if (unit != null)
            {
                _tokens.Add(new Token(TokenType.Length, sb.ToString(), startLine, startCol));
                return;
            }
        }

        _tokens.Add(new Token(TokenType.Number, sb.ToString(), startLine, startCol));
    }

    private void TokenizeIdentifierOrKeyword()
    {
        int startLine = _line;
        int startCol = _column;
        var sb = new StringBuilder();

        while (_pos < _source.Length && (char.IsLetterOrDigit(_source[_pos]) || _source[_pos] == '_' || _source[_pos] == '-'))
        {
            sb.Append(_source[_pos]);
            Advance();
        }

        var word = sb.ToString();

        // Check for booleans
        if (word == "true" || word == "false")
        {
            _tokens.Add(new Token(TokenType.Boolean, word, startLine, startCol));
            return;
        }

        // Check for document type values
        if (DocTypeValues.TryGetValue(word, out var docTypeToken))
        {
            _tokens.Add(new Token(docTypeToken, word, startLine, startCol));
            return;
        }

        // Check for keywords
        if (Keywords.TryGetValue(word, out var keywordToken))
        {
            _tokens.Add(new Token(keywordToken, word, startLine, startCol));
            return;
        }

        // Otherwise it's an identifier
        _tokens.Add(new Token(TokenType.Identifier, word, startLine, startCol));
    }

    private char Peek(int offset)
    {
        int idx = _pos + offset;
        return idx < _source.Length ? _source[idx] : '\0';
    }

    private void Advance()
    {
        if (_pos < _source.Length)
        {
            _pos++;
            _column++;
        }
    }

    private void SkipWhitespaceInLine()
    {
        while (_pos < _source.Length && (_source[_pos] == ' ' || _source[_pos] == '\t'))
            Advance();
    }

    private void SkipToNewLine()
    {
        while (_pos < _source.Length && _source[_pos] != '\n' && _source[_pos] != '\r')
            Advance();
    }

    private string ReadToEndOfLine()
    {
        var sb = new StringBuilder();
        while (_pos < _source.Length && _source[_pos] != '\n' && _source[_pos] != '\r')
        {
            sb.Append(_source[_pos]);
            Advance();
        }
        return sb.ToString();
    }

    private void ConsumeNewLine()
    {
        if (_pos < _source.Length && _source[_pos] == '\r')
            Advance();
        if (_pos < _source.Length && _source[_pos] == '\n')
            Advance();
        _line++;
        _column = 1;
    }

    private static bool IsHexDigit(char c) =>
        char.IsDigit(c) || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F');
}
