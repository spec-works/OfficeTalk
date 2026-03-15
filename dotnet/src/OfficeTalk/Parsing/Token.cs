namespace OfficeTalk.Parsing;

/// <summary>
/// Represents a single token produced by the OfficeTalk lexer.
/// </summary>
public class Token
{
    public TokenType Type { get; set; }
    public string Value { get; set; } = string.Empty;
    public int Line { get; set; }
    public int Column { get; set; }

    public Token() { }

    public Token(TokenType type, string value, int line, int column)
    {
        Type = type;
        Value = value;
        Line = line;
        Column = column;
    }

    public override string ToString() => $"{Type}({Value}) at {Line}:{Column}";
}

/// <summary>
/// All token types in the OfficeTalk grammar.
/// </summary>
public enum TokenType
{
    // Header
    Version,        // OFFICETALK/x.y

    // Keywords
    DocTypeKeyword, // DOCTYPE
    AT,
    EACH,
    SET,
    REPLACE,
    WITH,
    ALL,
    INSERT,
    BEFORE,
    AFTER,
    DELETE,
    APPEND,
    PREPEND,
    FORMAT,
    STYLE,
    PROPERTY,
    MERGE,
    CELLS,
    TO,
    ROW,
    COLUMN,
    SLIDE,
    SHEET,
    ADD,
    RENAME,
    DUPLICATE,
    COMMENT_KW,

    // Document type values
    Word,
    Excel,
    PowerPoint,

    // Literals
    String,         // "quoted string"
    Number,         // 42, 3.14, -7
    Boolean,        // true, false
    Color,          // #2B579A, #FF000080
    Length,         // 12pt, 1.5in, 2.54cm, 50%, 914400emu

    // Identifiers (for address segments, property names)
    Identifier,

    // Structural
    Slash,              // /
    LeftBracket,        // [
    RightBracket,       // ]
    Comma,              // ,
    Equals,             // =
    TildeEquals,        // ~=
    CaretEquals,        // ^=
    DollarEquals,       // $=
    AsteriskEquals,     // *=
    ContentBlockStart,  // <<<
    ContentBlockEnd,    // >>>
    ContentBlock,       // The content between <<< and >>>

    // Whitespace and structure
    Comment,            // # ...
    NewLine,
    EOF
}
