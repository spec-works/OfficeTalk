using DocumentFormat.OpenXml.Packaging;
using OfficeTalk.Addressing;
using OfficeTalk.Ast;

namespace OfficeTalk.Validation;

/// <summary>
/// Validates an OfficeTalk document against a target Office document.
/// Checks that addresses resolve, styles exist, operations are applicable, etc.
/// </summary>
public class SemanticValidator
{
    /// <summary>
    /// Validate the OfficeTalk document against the target Word document.
    /// </summary>
    public ValidationResult Validate(OfficeTalkDocument document, string targetPath)
    {
        var result = new ValidationResult();

        if (!File.Exists(targetPath))
        {
            result.Errors.Add(new ValidationDiagnostic(
                ValidationCategory.Syntax,
                $"Target document not found: '{targetPath}'."));
            return result;
        }

        if (document.DocType != DocType.Word)
        {
            // Only Word validation is currently supported
            return result;
        }

        using var wordDoc = WordprocessingDocument.Open(targetPath, false);
        var resolver = new WordAddressResolver(wordDoc);

        foreach (var block in document.OperationBlocks)
        {
            ValidateBlock(block, resolver, wordDoc, result);
        }

        return result;
    }

    private static void ValidateBlock(
        OperationBlock block,
        WordAddressResolver resolver,
        WordprocessingDocument wordDoc,
        ValidationResult result)
    {
        // Resolve address
        var elements = resolver.Resolve(block.Address);

        if (elements.Count == 0)
        {
            result.Errors.Add(new ValidationDiagnostic(
                ValidationCategory.AddressNotFound,
                $"Address '{block.Address.RawText}' did not match any elements.",
                block.Line));
            return;
        }

        if (!block.IsEach && elements.Count > 1)
        {
            result.Errors.Add(new ValidationDiagnostic(
                ValidationCategory.AddressAmbiguous,
                $"Address '{block.Address.RawText}' matched {elements.Count} elements. Use AT EACH for bulk operations.",
                block.Line));
        }

        // Validate STYLE references exist
        foreach (var op in block.Operations.OfType<StyleOperation>())
        {
            ValidateStyleExists(op, wordDoc, result);
        }

        // Validate REPLACE search text exists
        foreach (var op in block.Operations.OfType<ReplaceOperation>())
        {
            ValidateSearchTextExists(op, elements, result);
        }
    }

    private static void ValidateStyleExists(
        StyleOperation op,
        WordprocessingDocument wordDoc,
        ValidationResult result)
    {
        var stylesPart = wordDoc.MainDocumentPart?.StyleDefinitionsPart;
        if (stylesPart?.Styles == null) return;

        var styleExists = stylesPart.Styles
            .Elements<DocumentFormat.OpenXml.Wordprocessing.Style>()
            .Any(s => string.Equals(s.StyleName?.Val?.Value, op.StyleName, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(s.StyleId?.Value, op.StyleName, StringComparison.OrdinalIgnoreCase));

        if (!styleExists)
        {
            result.Errors.Add(new ValidationDiagnostic(
                ValidationCategory.MissingStyle,
                $"Style '{op.StyleName}' does not exist in the target document.",
                op.Line));
        }
    }

    private static void ValidateSearchTextExists(
        ReplaceOperation op,
        IReadOnlyList<DocumentFormat.OpenXml.OpenXmlElement> elements,
        ValidationResult result)
    {
        bool found = elements.Any(e =>
        {
            var text = e.InnerText;
            return text.Contains(op.Search, StringComparison.Ordinal);
        });

        if (!found)
        {
            result.Warnings.Add(new ValidationDiagnostic(
                ValidationCategory.SearchNotFound,
                $"REPLACE search text '{op.Search}' was not found in any matched element.",
                op.Line));
        }
    }
}
