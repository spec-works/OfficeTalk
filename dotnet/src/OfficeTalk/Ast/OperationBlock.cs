namespace OfficeTalk.Ast;

/// <summary>
/// An operation block: an address line (AT [EACH] address) followed by one or more operations.
/// </summary>
public class OperationBlock
{
    /// <summary>
    /// The target address for this block.
    /// </summary>
    public Address Address { get; set; } = new();

    /// <summary>
    /// Whether this is a bulk operation (AT EACH).
    /// </summary>
    public bool IsEach { get; set; }

    /// <summary>
    /// The operations to apply at the addressed element(s).
    /// </summary>
    public List<Operation> Operations { get; set; } = new();

    /// <summary>
    /// The source line number where this block begins.
    /// </summary>
    public int Line { get; set; }
}
