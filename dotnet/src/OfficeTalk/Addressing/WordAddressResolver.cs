using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeTalk.Ast;

namespace OfficeTalk.Addressing;

/// <summary>
/// Resolves OfficeTalk addresses against a Word (docx) document using the Open XML SDK.
/// </summary>
public class WordAddressResolver : IAddressResolver
{
    private readonly WordprocessingDocument _document;

    public WordAddressResolver(WordprocessingDocument document)
    {
        _document = document ?? throw new ArgumentNullException(nameof(document));
    }

    /// <inheritdoc/>
    public IReadOnlyList<OpenXmlElement> Resolve(Address address)
    {
        if (address.Segments.Count == 0)
            return Array.Empty<OpenXmlElement>();

        var body = _document.MainDocumentPart?.Document?.Body;
        if (body == null)
            return Array.Empty<OpenXmlElement>();

        IEnumerable<OpenXmlElement> current = new[] { body };

        foreach (var segment in address.Segments)
        {
            current = ResolveSegment(current, segment);
        }

        return current.ToList().AsReadOnly();
    }

    private IEnumerable<OpenXmlElement> ResolveSegment(IEnumerable<OpenXmlElement> parents, AddressSegment segment)
    {
        var results = new List<OpenXmlElement>();

        foreach (var parent in parents)
        {
            var candidates = GetCandidateElements(parent, segment.Identifier);
            var filtered = ApplyPredicates(candidates.ToList(), segment.Predicates);
            results.AddRange(filtered);
        }

        return results;
    }

    private IEnumerable<OpenXmlElement> GetCandidateElements(OpenXmlElement parent, string identifier)
    {
        return identifier.ToLowerInvariant() switch
        {
            "body" => parent is Body b ? new[] { b } : parent.Elements<Body>(),
            "paragraph" => GetParagraphs(parent),
            "heading" => GetHeadings(parent),
            "table" => parent.Elements<Table>(),
            "row" => parent.Elements<TableRow>(),
            "cell" => parent.Elements<TableCell>(),
            "run" => parent.Elements<Run>(),
            "bookmark" => throw new NotImplementedException("Bookmark address resolution is not yet implemented."),
            "list" => throw new NotImplementedException("List address resolution is not yet implemented."),
            "item" => throw new NotImplementedException("List item address resolution is not yet implemented."),
            "image" => throw new NotImplementedException("Image address resolution is not yet implemented."),
            "section" => throw new NotImplementedException("Section address resolution is not yet implemented."),
            "header" => throw new NotImplementedException("Header address resolution is not yet implemented."),
            "footer" => throw new NotImplementedException("Footer address resolution is not yet implemented."),
            "content-control" => throw new NotImplementedException("Content control address resolution is not yet implemented."),
            _ => Enumerable.Empty<OpenXmlElement>()
        };
    }

    private static IEnumerable<OpenXmlElement> GetParagraphs(OpenXmlElement parent)
    {
        // Return paragraphs that are NOT headings (no heading style)
        return parent.Elements<Paragraph>()
            .Where(p => !IsHeading(p));
    }

    private static IEnumerable<OpenXmlElement> GetHeadings(OpenXmlElement parent)
    {
        return parent.Elements<Paragraph>()
            .Where(IsHeading);
    }

    private static bool IsHeading(Paragraph paragraph)
    {
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId == null) return false;
        return styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)
            || styleId.StartsWith("heading", StringComparison.OrdinalIgnoreCase);
    }

    private static int GetHeadingLevel(Paragraph paragraph)
    {
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (styleId == null) return 0;

        // Extract level from "Heading1", "Heading2", etc.
        if (styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase) && styleId.Length > 7)
        {
            if (int.TryParse(styleId.AsSpan(7), out int level))
                return level;
        }

        return 0;
    }

    private static string GetParagraphText(OpenXmlElement element)
    {
        if (element is Paragraph p)
            return string.Concat(p.Elements<Run>().SelectMany(r => r.Elements<Text>()).Select(t => t.Text));
        return string.Empty;
    }

    private IReadOnlyList<OpenXmlElement> ApplyPredicates(List<OpenXmlElement> candidates, List<Ast.Predicate> predicates)
    {
        var results = candidates;

        foreach (var predicate in predicates)
        {
            results = predicate switch
            {
                PositionalPredicate pos => ApplyPositionalPredicate(results, pos),
                KeyValuePredicate kv => ApplyKeyValuePredicate(results, kv),
                BareStringPredicate bare => ApplyBareStringPredicate(results, bare),
                _ => results
            };
        }

        return results;
    }

    private static List<OpenXmlElement> ApplyPositionalPredicate(List<OpenXmlElement> candidates, PositionalPredicate pred)
    {
        if (pred.Position >= 1 && pred.Position <= candidates.Count)
            return new List<OpenXmlElement> { candidates[pred.Position - 1] };
        return new List<OpenXmlElement>();
    }

    private static List<OpenXmlElement> ApplyKeyValuePredicate(List<OpenXmlElement> candidates, KeyValuePredicate pred)
    {
        return candidates.Where(e => MatchesKeyValue(e, pred)).ToList();
    }

    private static List<OpenXmlElement> ApplyBareStringPredicate(List<OpenXmlElement> candidates, BareStringPredicate pred)
    {
        return candidates.Where(e => GetParagraphText(e) == pred.Value).ToList();
    }

    private static bool MatchesKeyValue(OpenXmlElement element, KeyValuePredicate pred)
    {
        switch (pred.Key.ToLowerInvariant())
        {
            case "level":
                if (element is Paragraph heading && IsHeading(heading))
                {
                    var level = GetHeadingLevel(heading);
                    return int.TryParse(pred.Value, out int targetLevel) && level == targetLevel;
                }
                return false;

            case "text":
                var text = GetParagraphText(element);
                return pred.Operator switch
                {
                    PredicateOperator.Equals => text == pred.Value,
                    PredicateOperator.AsteriskEquals => text.Contains(pred.Value, StringComparison.Ordinal),
                    PredicateOperator.CaretEquals => text.StartsWith(pred.Value, StringComparison.Ordinal),
                    PredicateOperator.DollarEquals => text.EndsWith(pred.Value, StringComparison.Ordinal),
                    PredicateOperator.TildeEquals => System.Text.RegularExpressions.Regex.IsMatch(text, pred.Value),
                    _ => false
                };

            case "caption":
                // Table caption matching — check preceding paragraph for caption text
                return false; // Stub for now

            case "tag":
            case "name":
                return false; // Stub for content-control/shape attributes

            default:
                return false;
        }
    }
}
