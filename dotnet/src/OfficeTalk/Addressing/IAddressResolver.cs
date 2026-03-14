using DocumentFormat.OpenXml;

namespace OfficeTalk.Addressing;

/// <summary>
/// Interface for resolving OfficeTalk addresses against a target document.
/// </summary>
public interface IAddressResolver
{
    /// <summary>
    /// Resolve an address to one or more elements in the target document.
    /// </summary>
    /// <param name="address">The OfficeTalk address to resolve.</param>
    /// <returns>The resolved elements.</returns>
    IReadOnlyList<OpenXmlElement> Resolve(Ast.Address address);
}
