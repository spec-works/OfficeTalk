namespace OfficeTalk.Validation;

/// <summary>
/// The result of validating an OfficeTalk document.
/// </summary>
public class ValidationResult
{
    /// <summary>
    /// Validation errors that prevent execution.
    /// </summary>
    public List<ValidationDiagnostic> Errors { get; set; } = new();

    /// <summary>
    /// Validation warnings that don't prevent execution.
    /// </summary>
    public List<ValidationDiagnostic> Warnings { get; set; } = new();

    /// <summary>
    /// Whether the document is valid (no errors).
    /// </summary>
    public bool IsValid => Errors.Count == 0;
}

/// <summary>
/// A single validation diagnostic (error or warning).
/// </summary>
public class ValidationDiagnostic
{
    public ValidationCategory Category { get; set; }
    public string Message { get; set; } = string.Empty;
    public int? Line { get; set; }
    public int? Column { get; set; }

    public ValidationDiagnostic() { }

    public ValidationDiagnostic(ValidationCategory category, string message, int? line = null, int? column = null)
    {
        Category = category;
        Message = message;
        Line = line;
        Column = column;
    }

    public override string ToString()
    {
        var loc = Line.HasValue ? $" at line {Line}" : "";
        return $"[{Category}]{loc}: {Message}";
    }
}

/// <summary>
/// Validation error/warning categories from the OfficeTalk specification.
/// </summary>
public enum ValidationCategory
{
    /// <summary>Grammar violation.</summary>
    Syntax,
    /// <summary>Address matches nothing in the target document.</summary>
    AddressNotFound,
    /// <summary>Address matches multiple elements when single expected.</summary>
    AddressAmbiguous,
    /// <summary>Operation not applicable to the target element type.</summary>
    InvalidOperation,
    /// <summary>Property value has wrong type or is out of range.</summary>
    InvalidValue,
    /// <summary>Referenced style does not exist in the target document.</summary>
    MissingStyle,
    /// <summary>REPLACE search text not found in target content.</summary>
    SearchNotFound,
    /// <summary>Conflicting operations on the same element.</summary>
    Conflict,
    /// <summary>Invalid structural operation.</summary>
    Structural,
    /// <summary>FORMAT overrides STYLE in the same block.</summary>
    StyleOverride,
    /// <summary>Operation has no effect.</summary>
    RedundantOp,
    /// <summary>Use of deprecated syntax.</summary>
    Deprecated,
    /// <summary>Document mixes INSPECT and write operations.</summary>
    MixedOperations
}
