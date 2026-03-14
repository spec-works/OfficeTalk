using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FluentAssertions;
using OfficeTalk.Addressing;
using OfficeTalk.Ast;

namespace OfficeTalk.Tests.Addressing;

public class WordAddressResolverTests : IDisposable
{
    private readonly MemoryStream _stream;
    private readonly WordprocessingDocument _document;

    public WordAddressResolverTests()
    {
        _stream = new MemoryStream();
        _document = WordprocessingDocument.Create(_stream, WordprocessingDocumentType.Document);
        var mainPart = _document.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var body = mainPart.Document.Body!;

        // Add Heading1
        body.AppendChild(CreateHeading("Introduction", 1));

        // Add 3 paragraphs
        body.AppendChild(CreateParagraph("First paragraph of the document."));
        body.AppendChild(CreateParagraph("Second paragraph with some content."));
        body.AppendChild(CreateParagraph("Third paragraph about the conclusion."));

        // Add Heading2
        body.AppendChild(CreateHeading("Methods", 2));

        // Add another paragraph
        body.AppendChild(CreateParagraph("Methods section content here."));

        // Add a table with 2 rows
        var table = new Table();
        var row1 = new TableRow(new TableCell(new Paragraph(new Run(new Text("Header 1")))));
        var row2 = new TableRow(new TableCell(new Paragraph(new Run(new Text("Data 1")))));
        table.AppendChild(row1);
        table.AppendChild(row2);
        body.AppendChild(table);

        mainPart.Document.Save();
    }

    public void Dispose()
    {
        _document.Dispose();
        _stream.Dispose();
    }

    [Fact]
    public void Resolve_BodyParagraphPositional_ReturnsCorrectParagraph()
    {
        var resolver = new WordAddressResolver(_document);
        var address = CreateAddress(("body", null), ("paragraph", new PositionalPredicate(2)));

        var results = resolver.Resolve(address);

        results.Should().HaveCount(1);
        results[0].InnerText.Should().Be("Second paragraph with some content.");
    }

    [Fact]
    public void Resolve_BodyHeadingByLevel_ReturnsHeadings()
    {
        var resolver = new WordAddressResolver(_document);
        var address = CreateAddress(
            ("body", null),
            ("heading", new KeyValuePredicate("level", PredicateOperator.Equals, "1")));

        var results = resolver.Resolve(address);

        results.Should().HaveCount(1);
        results[0].InnerText.Should().Be("Introduction");
    }

    [Fact]
    public void Resolve_HeadingByLevelAndText_ReturnsSpecificHeading()
    {
        var resolver = new WordAddressResolver(_document);

        var segment = new AddressSegment
        {
            Identifier = "heading",
            Predicates = new List<Ast.Predicate>
            {
                new KeyValuePredicate("level", PredicateOperator.Equals, "2"),
                new KeyValuePredicate("text", PredicateOperator.Equals, "Methods")
            }
        };

        var address = new Address
        {
            Segments = new List<AddressSegment>
            {
                new AddressSegment { Identifier = "body" },
                segment
            }
        };

        var results = resolver.Resolve(address);

        results.Should().HaveCount(1);
        results[0].InnerText.Should().Be("Methods");
    }

    [Fact]
    public void Resolve_ParagraphWithTextContains_MatchesCorrectParagraph()
    {
        var resolver = new WordAddressResolver(_document);
        var address = CreateAddress(
            ("body", null),
            ("paragraph", new KeyValuePredicate("text", PredicateOperator.AsteriskEquals, "conclusion")));

        var results = resolver.Resolve(address);

        results.Should().HaveCount(1);
        results[0].InnerText.Should().Contain("conclusion");
    }

    [Fact]
    public void Resolve_AllParagraphs_ReturnsNonHeadingParagraphs()
    {
        var resolver = new WordAddressResolver(_document);
        var address = CreateAddress(("body", null), ("paragraph", null));

        var results = resolver.Resolve(address);

        // Should return all non-heading paragraphs (4 regular + those inside table cells)
        results.Should().HaveCountGreaterOrEqualTo(4);
    }

    [Fact]
    public void Resolve_Table_ReturnsTableElement()
    {
        var resolver = new WordAddressResolver(_document);
        var address = CreateAddress(("body", null), ("table", new PositionalPredicate(1)));

        var results = resolver.Resolve(address);

        results.Should().HaveCount(1);
        results[0].Should().BeOfType<Table>();
    }

    [Fact]
    public void Resolve_TableRow_ReturnsRows()
    {
        var resolver = new WordAddressResolver(_document);
        var address = CreateAddress(
            ("body", null),
            ("table", new PositionalPredicate(1)),
            ("row", new PositionalPredicate(2)));

        var results = resolver.Resolve(address);

        results.Should().HaveCount(1);
        results[0].Should().BeOfType<TableRow>();
    }

    [Fact]
    public void Resolve_NonExistentAddress_ReturnsEmpty()
    {
        var resolver = new WordAddressResolver(_document);
        var address = CreateAddress(("body", null), ("paragraph", new PositionalPredicate(999)));

        var results = resolver.Resolve(address);

        results.Should().BeEmpty();
    }

    [Fact]
    public void Resolve_ParagraphTextStartsWith_MatchesCorrectly()
    {
        var resolver = new WordAddressResolver(_document);
        var address = CreateAddress(
            ("body", null),
            ("paragraph", new KeyValuePredicate("text", PredicateOperator.CaretEquals, "First")));

        var results = resolver.Resolve(address);

        results.Should().HaveCount(1);
        results[0].InnerText.Should().StartWith("First");
    }

    [Fact]
    public void Resolve_ParagraphTextEndsWith_MatchesCorrectly()
    {
        var resolver = new WordAddressResolver(_document);
        var address = CreateAddress(
            ("body", null),
            ("paragraph", new KeyValuePredicate("text", PredicateOperator.DollarEquals, "here.")));

        var results = resolver.Resolve(address);

        results.Should().HaveCount(1);
        results[0].InnerText.Should().EndWith("here.");
    }

    private static Address CreateAddress(params (string identifier, Ast.Predicate? predicate)[] segments)
    {
        var address = new Address();
        foreach (var (identifier, predicate) in segments)
        {
            var segment = new AddressSegment { Identifier = identifier };
            if (predicate != null)
                segment.Predicates.Add(predicate);
            address.Segments.Add(segment);
        }
        return address;
    }

    private static Paragraph CreateParagraph(string text)
    {
        return new Paragraph(new Run(new Text(text)));
    }

    private static Paragraph CreateHeading(string text, int level)
    {
        var paragraph = new Paragraph(
            new ParagraphProperties(
                new ParagraphStyleId { Val = $"Heading{level}" }),
            new Run(new Text(text)));
        return paragraph;
    }
}
