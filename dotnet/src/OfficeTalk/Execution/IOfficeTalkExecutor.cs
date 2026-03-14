using OfficeTalk.Ast;

namespace OfficeTalk.Execution;

/// <summary>
/// Interface for executing OfficeTalk operations against a target document.
/// </summary>
public interface IOfficeTalkExecutor
{
    /// <summary>
    /// Execute all operation blocks in the document against the target.
    /// </summary>
    /// <param name="document">The parsed OfficeTalk document.</param>
    /// <param name="targetPath">Path to the target Office document.</param>
    /// <param name="outputPath">Path for the output document (null = modify in place).</param>
    void Execute(OfficeTalkDocument document, string targetPath, string? outputPath = null);
}
