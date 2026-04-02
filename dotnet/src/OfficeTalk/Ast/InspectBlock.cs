namespace OfficeTalk.Ast;

/// <summary>
/// An inspect block: an INSPECT address line followed by optional modifiers.
/// </summary>
public class InspectBlock
{
    /// <summary>
    /// The target address to inspect.
    /// </summary>
    public Address Address { get; set; } = new();

    /// <summary>
    /// Number of child element levels to include (default 0).
    /// </summary>
    public int Depth { get; set; }

    /// <summary>
    /// Which detail layers to include (content, properties, or both).
    /// </summary>
    public List<IncludeLayer> Include { get; set; } = new();

    /// <summary>
    /// Number of sibling elements before and after to include (default 0).
    /// </summary>
    public int Context { get; set; }

    /// <summary>
    /// The source line number where this block begins.
    /// </summary>
    public int Line { get; set; }
}

/// <summary>
/// Detail layers that can be requested in an INSPECT operation.
/// </summary>
public enum IncludeLayer
{
    /// <summary>Textual content of the element.</summary>
    Content,
    /// <summary>Formatting and metadata properties.</summary>
    Properties
}
