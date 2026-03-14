namespace OfficeTalk.Ast;

/// <summary>
/// An address composed of path segments separated by slashes.
/// Example: body/heading[level=2, text="Methods"]
/// </summary>
public class Address
{
    /// <summary>
    /// The segments that make up this address path.
    /// </summary>
    public List<AddressSegment> Segments { get; set; } = new();

    /// <summary>
    /// The raw text of the address for error reporting.
    /// </summary>
    public string RawText { get; set; } = string.Empty;

    public override string ToString() =>
        string.Join("/", Segments.Select(s => s.ToString()));
}
