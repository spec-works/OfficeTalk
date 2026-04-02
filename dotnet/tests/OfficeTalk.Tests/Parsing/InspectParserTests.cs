using FluentAssertions;
using OfficeTalk.Ast;
using OfficeTalk.Parsing;
using OfficeTalk.Validation;

namespace OfficeTalk.Tests.Parsing;

public class InspectParserTests
{
    private static OfficeTalkDocument Parse(string input)
    {
        var lexer = new OfficeTalkLexer(input);
        var tokens = lexer.Tokenize();
        var parser = new OfficeTalkParser(tokens);
        return parser.Parse();
    }

    // ─── Lexer token recognition ─────────────────────────────────

    [Theory]
    [InlineData("INSPECT", TokenType.INSPECT)]
    [InlineData("DEPTH", TokenType.DEPTH)]
    [InlineData("INCLUDE", TokenType.INCLUDE)]
    [InlineData("CONTEXT", TokenType.CONTEXT)]
    public void Lexer_InspectKeywords_RecognizedCorrectly(string keyword, TokenType expected)
    {
        var lexer = new OfficeTalkLexer($"OFFICETALK/1.0\nDOCTYPE word\n\n{keyword}\n");
        var tokens = lexer.Tokenize();

        tokens.Should().Contain(t => t.Type == expected);
    }

    [Fact]
    public void Lexer_InspectKeywords_CaseInsensitive()
    {
        var lexer = new OfficeTalkLexer("OFFICETALK/1.0\nDOCTYPE word\n\ninspect\ndepth\ninclude\ncontext\n");
        var tokens = lexer.Tokenize();

        tokens.Should().Contain(t => t.Type == TokenType.INSPECT);
        tokens.Should().Contain(t => t.Type == TokenType.DEPTH);
        tokens.Should().Contain(t => t.Type == TokenType.INCLUDE);
        tokens.Should().Contain(t => t.Type == TokenType.CONTEXT);
    }

    // ─── Basic INSPECT parsing ───────────────────────────────────

    [Fact]
    public void Parse_SimpleInspect_ProducesInspectBlock()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading\n");

        doc.InspectBlocks.Should().HaveCount(1);
        doc.OperationBlocks.Should().BeEmpty();

        var block = doc.InspectBlocks[0];
        block.Address.Segments.Should().HaveCount(2);
        block.Address.Segments[0].Identifier.Should().Be("body");
        block.Address.Segments[1].Identifier.Should().Be("heading");
        block.Depth.Should().Be(0);
        block.Include.Should().BeEmpty();
        block.Context.Should().Be(0);
    }

    [Fact]
    public void Parse_InspectWithPositionalPredicate_ParsedCorrectly()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE excel\n\nINSPECT sheet[1]\n");

        doc.InspectBlocks.Should().HaveCount(1);
        var seg = doc.InspectBlocks[0].Address.Segments[0];
        seg.Identifier.Should().Be("sheet");
        seg.Predicates.Should().HaveCount(1);
        seg.Predicates[0].Should().BeOfType<PositionalPredicate>()
            .Which.Position.Should().Be(1);
    }

    [Fact]
    public void Parse_InspectWithBareStringPredicate_ParsedCorrectly()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE excel\n\nINSPECT sheet[\"Q1 Budget\"]\n");

        var seg = doc.InspectBlocks[0].Address.Segments[0];
        seg.Predicates[0].Should().BeOfType<BareStringPredicate>()
            .Which.Value.Should().Be("Q1 Budget");
    }

    [Fact]
    public void Parse_InspectWithKeyValuePredicate_ParsedCorrectly()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading[text=\"Conclusion\"]\n");

        var seg = doc.InspectBlocks[0].Address.Segments[1];
        seg.Predicates[0].Should().BeOfType<KeyValuePredicate>();
        var pred = (KeyValuePredicate)seg.Predicates[0];
        pred.Key.Should().Be("text");
        pred.Operator.Should().Be(PredicateOperator.Equals);
        pred.Value.Should().Be("Conclusion");
    }

    [Fact]
    public void Parse_InspectDeepAddress_ParsedCorrectly()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE excel\n\nINSPECT sheet[\"Q1 Budget\"]/D2\n");

        var segs = doc.InspectBlocks[0].Address.Segments;
        segs.Should().HaveCount(2);
        segs[0].Identifier.Should().Be("sheet");
        segs[1].Identifier.Should().Be("D2");
    }

    // ─── Modifier parsing ────────────────────────────────────────

    [Fact]
    public void Parse_InspectWithDepth_SetsDepthValue()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE excel\n\nINSPECT sheet[1]\n  DEPTH 2\n");

        doc.InspectBlocks[0].Depth.Should().Be(2);
    }

    [Fact]
    public void Parse_InspectWithIncludeContent_SetsContentLayer()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading\n  INCLUDE content\n");

        doc.InspectBlocks[0].Include.Should().ContainSingle()
            .Which.Should().Be(IncludeLayer.Content);
    }

    [Fact]
    public void Parse_InspectWithIncludeProperties_SetsPropertiesLayer()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/paragraph[1]\n  INCLUDE properties\n");

        doc.InspectBlocks[0].Include.Should().ContainSingle()
            .Which.Should().Be(IncludeLayer.Properties);
    }

    [Fact]
    public void Parse_InspectWithIncludeBoth_SetsBothLayers()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE excel\n\nINSPECT sheet[\"Q1 Budget\"]/D2\n  INCLUDE content, properties\n");

        var block = doc.InspectBlocks[0];
        block.Include.Should().HaveCount(2);
        block.Include.Should().Contain(IncludeLayer.Content);
        block.Include.Should().Contain(IncludeLayer.Properties);
    }

    [Fact]
    public void Parse_InspectWithContext_SetsContextValue()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE excel\n\nINSPECT sheet[\"Q1 Budget\"]/D2\n  CONTEXT 3\n");

        doc.InspectBlocks[0].Context.Should().Be(3);
    }

    [Fact]
    public void Parse_InspectWithAllModifiers_AllSet()
    {
        var input = @"OFFICETALK/1.0
DOCTYPE word

INSPECT body/table[1]
  DEPTH 2
  INCLUDE content, properties
  CONTEXT 1
";
        var doc = Parse(input);

        var block = doc.InspectBlocks[0];
        block.Depth.Should().Be(2);
        block.Include.Should().HaveCount(2);
        block.Include.Should().Contain(IncludeLayer.Content);
        block.Include.Should().Contain(IncludeLayer.Properties);
        block.Context.Should().Be(1);
    }

    [Fact]
    public void Parse_InspectModifiersAnyOrder_AllSet()
    {
        var input = @"OFFICETALK/1.0
DOCTYPE word

INSPECT body/heading
  CONTEXT 2
  INCLUDE content
  DEPTH 1
";
        var doc = Parse(input);

        var block = doc.InspectBlocks[0];
        block.Depth.Should().Be(1);
        block.Include.Should().ContainSingle().Which.Should().Be(IncludeLayer.Content);
        block.Context.Should().Be(2);
    }

    // ─── Multiple INSPECT blocks ─────────────────────────────────

    [Fact]
    public void Parse_MultipleInspectBlocks_AllParsed()
    {
        var input = @"OFFICETALK/1.0
DOCTYPE excel

INSPECT sheet[1]

INSPECT sheet[""Q1 Budget""]
  DEPTH 1
  INCLUDE content

INSPECT sheet[""Q1 Budget""]/D2
  INCLUDE content, properties
";
        var doc = Parse(input);

        doc.InspectBlocks.Should().HaveCount(3);
        doc.InspectBlocks[0].Depth.Should().Be(0);
        doc.InspectBlocks[0].Include.Should().BeEmpty();
        doc.InspectBlocks[1].Depth.Should().Be(1);
        doc.InspectBlocks[1].Include.Should().ContainSingle();
        doc.InspectBlocks[2].Include.Should().HaveCount(2);
    }

    // ─── Comments within INSPECT blocks ──────────────────────────

    [Fact]
    public void Parse_InspectWithComments_CommentsIgnored()
    {
        var input = @"OFFICETALK/1.0
DOCTYPE word

# Inspect the document outline
INSPECT body/heading
  # Include text content
  INCLUDE content
";
        var doc = Parse(input);

        doc.InspectBlocks.Should().HaveCount(1);
        doc.InspectBlocks[0].Include.Should().ContainSingle()
            .Which.Should().Be(IncludeLayer.Content);
    }

    // ─── All three doctypes ──────────────────────────────────────

    [Theory]
    [InlineData("word", "body/heading")]
    [InlineData("excel", "sheet[1]")]
    [InlineData("powerpoint", "slide[1]")]
    public void Parse_InspectAllDocTypes_Recognized(string docType, string address)
    {
        var doc = Parse($"OFFICETALK/1.0\nDOCTYPE {docType}\n\nINSPECT {address}\n");

        doc.InspectBlocks.Should().HaveCount(1);
        doc.Errors.Should().BeEmpty();
    }

    // ─── Line number tracking ────────────────────────────────────

    [Fact]
    public void Parse_InspectBlock_TracksLineNumber()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading\n  DEPTH 1\n";
        var doc = Parse(input);

        doc.InspectBlocks[0].Line.Should().Be(4);
    }

    // ─── Validation: mixed operations ────────────────────────────

    [Fact]
    public void Validate_MixedInspectAndAt_ProducesError()
    {
        var input = @"OFFICETALK/1.0
DOCTYPE word

INSPECT body/heading

AT body/paragraph[1]
SET ""hello""
";
        var doc = Parse(input);

        var validator = new SyntacticValidator();
        var result = validator.Validate(doc);

        result.IsValid.Should().BeFalse();
        result.Errors.Should().Contain(e => e.Category == ValidationCategory.MixedOperations);
    }

    [Fact]
    public void Validate_MixedInspectAndProperty_ProducesError()
    {
        var input = @"OFFICETALK/1.0
DOCTYPE word

INSPECT body/heading

PROPERTY title=""test""
";
        var doc = Parse(input);

        var validator = new SyntacticValidator();
        var result = validator.Validate(doc);

        result.IsValid.Should().BeFalse();
        result.Errors.Should().Contain(e => e.Category == ValidationCategory.MixedOperations);
    }

    [Fact]
    public void Validate_InspectOnly_IsValid()
    {
        var input = @"OFFICETALK/1.0
DOCTYPE word

INSPECT body/heading
  INCLUDE content
";
        var doc = Parse(input);

        var validator = new SyntacticValidator();
        var result = validator.Validate(doc);

        result.IsValid.Should().BeTrue();
    }

    [Fact]
    public void Validate_WriteOnly_IsValid()
    {
        var input = @"OFFICETALK/1.0
DOCTYPE word

AT body/paragraph[1]
SET ""hello""
";
        var doc = Parse(input);

        var validator = new SyntacticValidator();
        var result = validator.Validate(doc);

        result.IsValid.Should().BeTrue();
    }

    // ─── Error handling ──────────────────────────────────────────

    [Fact]
    public void Parse_InspectWithInvalidIncludeLayer_ProducesError()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading\n  INCLUDE invalid\n");

        doc.Errors.Should().Contain(e => e.Message.Contains("Unknown INCLUDE layer"));
    }

    [Fact]
    public void Parse_DepthWithoutNumber_ProducesError()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading\n  DEPTH\n");

        doc.Errors.Should().Contain(e => e.Message.Contains("Expected integer after DEPTH"));
    }

    [Fact]
    public void Parse_ContextWithoutNumber_ProducesError()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading\n  CONTEXT\n");

        doc.Errors.Should().Contain(e => e.Message.Contains("Expected integer after CONTEXT"));
    }

    [Fact]
    public void Parse_IncludeWithoutLayers_ProducesError()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading\n  INCLUDE\n");

        doc.Errors.Should().Contain(e => e.Message.Contains("Expected at least one layer"));
    }

    // ─── Full spec examples ──────────────────────────────────────

    [Fact]
    public void Parse_ExcelInspectExample_ParsedCorrectly()
    {
        var input = @"OFFICETALK/1.0
DOCTYPE excel

# Discover what sheets exist (addressing only)
INSPECT sheet[1]

# See all rows in the Q1 Budget sheet with their values
INSPECT sheet[""Q1 Budget""]
  DEPTH 1
  INCLUDE content

# Get formatting details for a specific cell
INSPECT sheet[""Q1 Budget""]/D2
  INCLUDE content, properties

# See a cell with its neighbors
INSPECT sheet[""Q1 Budget""]/D2
  INCLUDE content
  CONTEXT 2
";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        doc.InspectBlocks.Should().HaveCount(4);

        // Block 1: addressing only
        doc.InspectBlocks[0].Depth.Should().Be(0);
        doc.InspectBlocks[0].Include.Should().BeEmpty();
        doc.InspectBlocks[0].Context.Should().Be(0);

        // Block 2: depth + content
        doc.InspectBlocks[1].Depth.Should().Be(1);
        doc.InspectBlocks[1].Include.Should().ContainSingle()
            .Which.Should().Be(IncludeLayer.Content);

        // Block 3: content + properties
        doc.InspectBlocks[2].Include.Should().HaveCount(2);

        // Block 4: content + context
        doc.InspectBlocks[3].Include.Should().ContainSingle();
        doc.InspectBlocks[3].Context.Should().Be(2);
    }

    [Fact]
    public void Parse_WordInspectExample_ParsedCorrectly()
    {
        var input = @"OFFICETALK/1.0
DOCTYPE word

# Get the document outline — all headings
INSPECT body/heading
  INCLUDE content

# Inspect a specific heading with its surrounding paragraphs
INSPECT body/heading[text=""Conclusion""]
  INCLUDE content
  CONTEXT 3

# See a table's full structure with formatting
INSPECT body/table[1]
  DEPTH 2
  INCLUDE content, properties
";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        doc.InspectBlocks.Should().HaveCount(3);

        doc.InspectBlocks[1].Context.Should().Be(3);
        doc.InspectBlocks[2].Depth.Should().Be(2);
        doc.InspectBlocks[2].Include.Should().HaveCount(2);
    }

    [Fact]
    public void Parse_PowerPointInspectExample_ParsedCorrectly()
    {
        var input = @"OFFICETALK/1.0
DOCTYPE powerpoint

# Outline all slides (titles only)
INSPECT slide
  DEPTH 1
  INCLUDE content

# Full detail on slide 3 including comments
INSPECT slide[3]
  DEPTH 1
  INCLUDE content, properties
";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        doc.InspectBlocks.Should().HaveCount(2);

        doc.InspectBlocks[0].Address.Segments[0].Identifier.Should().Be("slide");
        doc.InspectBlocks[0].Address.Segments[0].Predicates.Should().BeEmpty();
        doc.InspectBlocks[1].Address.Segments[0].Predicates.Should().HaveCount(1);
    }

    // ─── Edge cases ──────────────────────────────────────────────

    [Fact]
    public void Parse_InspectDepthZero_ExplicitlySetToZero()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading\n  DEPTH 0\n");

        doc.InspectBlocks[0].Depth.Should().Be(0);
        doc.Errors.Should().BeEmpty();
    }

    [Fact]
    public void Parse_InspectWithDuplicateIncludeLayers_Deduplicated()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading\n  INCLUDE content, content\n");

        doc.InspectBlocks[0].Include.Should().ContainSingle()
            .Which.Should().Be(IncludeLayer.Content);
    }

    [Fact]
    public void Parse_InspectWithContainsOperator_ParsedCorrectly()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/paragraph[text*=\"conclusion\"]\n  INCLUDE content\n");

        var pred = doc.InspectBlocks[0].Address.Segments[1].Predicates[0]
            .Should().BeOfType<KeyValuePredicate>().Subject;
        pred.Key.Should().Be("text");
        pred.Operator.Should().Be(PredicateOperator.AsteriskEquals);
        pred.Value.Should().Be("conclusion");
    }

    [Fact]
    public void Parse_InspectAddressWithNoPredicates_MatchesAllElements()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading\n");

        var seg = doc.InspectBlocks[0].Address.Segments[1];
        seg.Identifier.Should().Be("heading");
        seg.Predicates.Should().BeEmpty();
    }

    [Fact]
    public void Parse_InspectLargeDepth_ParsedCorrectly()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body\n  DEPTH 10\n");

        doc.InspectBlocks[0].Depth.Should().Be(10);
        doc.Errors.Should().BeEmpty();
    }

    [Fact]
    public void Parse_InspectLargeContext_ParsedCorrectly()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/paragraph[1]\n  CONTEXT 100\n");

        doc.InspectBlocks[0].Context.Should().Be(100);
        doc.Errors.Should().BeEmpty();
    }

    [Fact]
    public void Parse_InspectStartsWithOperator_ParsedCorrectly()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading[text^=\"Chapter\"]\n");

        var pred = doc.InspectBlocks[0].Address.Segments[1].Predicates[0]
            .Should().BeOfType<KeyValuePredicate>().Subject;
        pred.Operator.Should().Be(PredicateOperator.CaretEquals);
        pred.Value.Should().Be("Chapter");
    }

    [Fact]
    public void Parse_InspectEndsWithOperator_ParsedCorrectly()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading[text$=\"Summary\"]\n");

        var pred = doc.InspectBlocks[0].Address.Segments[1].Predicates[0]
            .Should().BeOfType<KeyValuePredicate>().Subject;
        pred.Operator.Should().Be(PredicateOperator.DollarEquals);
        pred.Value.Should().Be("Summary");
    }

    [Fact]
    public void Parse_InspectMultiplePredicates_ParsedCorrectly()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading[level=\"1\", text^=\"Ch\"]\n");

        var preds = doc.InspectBlocks[0].Address.Segments[1].Predicates;
        preds.Should().HaveCount(2);
        preds[0].Should().BeOfType<KeyValuePredicate>();
        preds[1].Should().BeOfType<KeyValuePredicate>();
    }

    [Fact]
    public void Parse_InspectDeepNestedAddress_ParsedCorrectly()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/table[1]/row[2]/cell[3]\n");

        doc.Errors.Should().BeEmpty();
        doc.InspectBlocks[0].Address.Segments.Should().HaveCount(4);
        doc.InspectBlocks[0].Address.Segments[3].Identifier.Should().Be("cell");
    }

    [Fact]
    public void Parse_InspectModifiersReversed_AllSet()
    {
        var doc = Parse(@"OFFICETALK/1.0
DOCTYPE word

INSPECT body/heading
  CONTEXT 2
  INCLUDE properties
  DEPTH 3
");
        doc.Errors.Should().BeEmpty();
        doc.InspectBlocks[0].Context.Should().Be(2);
        doc.InspectBlocks[0].Include.Should().Contain(IncludeLayer.Properties);
        doc.InspectBlocks[0].Depth.Should().Be(3);
    }

    [Fact]
    public void Parse_InspectOnlyIncludeContentComma_SingleLayer()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/paragraph\n  INCLUDE content\n");

        doc.InspectBlocks[0].Include.Should().ContainSingle()
            .Which.Should().Be(IncludeLayer.Content);
    }

    [Fact]
    public void Parse_InspectOnlyIncludeProperties_SingleLayer()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/paragraph\n  INCLUDE properties\n");

        doc.InspectBlocks[0].Include.Should().ContainSingle()
            .Which.Should().Be(IncludeLayer.Properties);
    }

    [Fact]
    public void Parse_InspectNoModifiers_DefaultValues()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading\n");

        doc.InspectBlocks[0].Depth.Should().Be(0);
        doc.InspectBlocks[0].Include.Should().BeEmpty();
        doc.InspectBlocks[0].Context.Should().Be(0);
    }

    [Fact]
    public void Parse_InspectFollowedByBlankLines_ParsedCorrectly()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nINSPECT body/heading\n  INCLUDE content\n\n\n");

        doc.Errors.Should().BeEmpty();
        doc.InspectBlocks.Should().HaveCount(1);
    }

    [Fact]
    public void Parse_ManyInspectBlocks_AllParsed()
    {
        var sb = new System.Text.StringBuilder();
        sb.AppendLine("OFFICETALK/1.0");
        sb.AppendLine("DOCTYPE word");
        sb.AppendLine();
        for (int i = 1; i <= 10; i++)
        {
            sb.AppendLine($"INSPECT body/paragraph[{i}]");
            sb.AppendLine("  INCLUDE content");
            sb.AppendLine();
        }
        var doc = Parse(sb.ToString());

        doc.Errors.Should().BeEmpty();
        doc.InspectBlocks.Should().HaveCount(10);
        for (int i = 0; i < 10; i++)
        {
            var pred = doc.InspectBlocks[i].Address.Segments[1].Predicates[0]
                .Should().BeOfType<PositionalPredicate>().Subject;
            pred.Position.Should().Be(i + 1);
        }
    }

    // ─── Validation edge cases ──────────────────────────────────

    [Fact]
    public void Validate_InspectWithProperty_ProducesError()
    {
        var doc = Parse(@"OFFICETALK/1.0
DOCTYPE word

INSPECT body/heading

PROPERTY title = ""My Doc""
");
        var validator = new SyntacticValidator();
        var result = validator.Validate(doc);

        result.IsValid.Should().BeFalse();
        result.Errors.Should().Contain(e =>
            e.Category == ValidationCategory.MixedOperations);
    }

    [Fact]
    public void Validate_MultipleInspectBlocks_IsValid()
    {
        var doc = Parse(@"OFFICETALK/1.0
DOCTYPE word

INSPECT body/heading
  INCLUDE content

INSPECT body/table
  DEPTH 2
");
        var validator = new SyntacticValidator();
        var result = validator.Validate(doc);

        result.IsValid.Should().BeTrue();
    }
}
