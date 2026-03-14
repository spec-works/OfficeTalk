namespace OfficeTalk.Ast;

/// <summary>
/// A single segment in an address path, with optional predicates.
/// Example: heading[level=2, text="Methods"]
/// </summary>
public class AddressSegment
{
    /// <summary>
    /// The segment identifier (e.g., "body", "paragraph", "heading", "table").
    /// </summary>
    public string Identifier { get; set; } = string.Empty;

    /// <summary>
    /// Predicates that filter elements at this segment.
    /// </summary>
    public List<Predicate> Predicates { get; set; } = new();

    public override string ToString()
    {
        if (Predicates.Count == 0)
            return Identifier;

        var preds = string.Join(", ", Predicates.Select(p => p.ToString()));
        return $"{Identifier}[{preds}]";
    }
}
