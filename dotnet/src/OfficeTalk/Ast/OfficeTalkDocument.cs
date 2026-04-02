namespace OfficeTalk.Ast;

/// <summary>
/// Root AST node representing a complete OfficeTalk document.
/// </summary>
public class OfficeTalkDocument
{
    /// <summary>
    /// The OfficeTalk version (e.g., "1.0").
    /// </summary>
    public string Version { get; set; } = string.Empty;

    /// <summary>
    /// The target document type.
    /// </summary>
    public DocType DocType { get; set; }

    /// <summary>
    /// Operation blocks (AT address + operations).
    /// </summary>
    public List<OperationBlock> OperationBlocks { get; set; } = new();

    /// <summary>
    /// Inspect blocks (INSPECT address + modifiers).
    /// </summary>
    public List<InspectBlock> InspectBlocks { get; set; } = new();

    /// <summary>
    /// Document-level property settings.
    /// </summary>
    public List<PropertySetting> PropertySettings { get; set; } = new();

    /// <summary>
    /// Parse errors collected during parsing.
    /// </summary>
    public List<ParseError> Errors { get; set; } = new();
}

/// <summary>
/// The target document type for an OfficeTalk document.
/// </summary>
public enum DocType
{
    Word,
    Excel,
    PowerPoint
}

/// <summary>
/// A document-level property setting (e.g., PROPERTY title="...").
/// </summary>
public class PropertySetting
{
    public string Name { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
    public int Line { get; set; }
}

/// <summary>
/// A parse error with location information.
/// </summary>
public class ParseError
{
    public string Message { get; set; } = string.Empty;
    public int Line { get; set; }
    public int Column { get; set; }

    public ParseError() { }

    public ParseError(string message, int line, int column)
    {
        Message = message;
        Line = line;
        Column = column;
    }

    public override string ToString() => $"Line {Line}, Col {Column}: {Message}";
}
