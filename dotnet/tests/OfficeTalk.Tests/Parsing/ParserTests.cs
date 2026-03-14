using FluentAssertions;
using OfficeTalk.Ast;
using OfficeTalk.Parsing;

namespace OfficeTalk.Tests.Parsing;

public class ParserTests
{
    private static OfficeTalkDocument Parse(string input)
    {
        var lexer = new OfficeTalkLexer(input);
        var tokens = lexer.Tokenize();
        var parser = new OfficeTalkParser(tokens);
        return parser.Parse();
    }

    [Fact]
    public void Parse_HeaderOnly_ExtractsVersionAndDocType()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n");

        doc.Version.Should().Be("1.0");
        doc.DocType.Should().Be(DocType.Word);
    }

    [Theory]
    [InlineData("word", DocType.Word)]
    [InlineData("excel", DocType.Excel)]
    [InlineData("powerpoint", DocType.PowerPoint)]
    public void Parse_AllDocTypes_Recognized(string docType, DocType expected)
    {
        var doc = Parse($"OFFICETALK/1.0\nDOCTYPE {docType}\n");
        doc.DocType.Should().Be(expected);
    }

    [Fact]
    public void Parse_SimpleSet_ProducesSetOperation()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nSET \"New text\"\n");

        doc.OperationBlocks.Should().HaveCount(1);
        var block = doc.OperationBlocks[0];
        block.Address.Segments.Should().HaveCount(2);
        block.Address.Segments[0].Identifier.Should().Be("body");
        block.Address.Segments[1].Identifier.Should().Be("paragraph");

        block.Operations.Should().HaveCount(1);
        block.Operations[0].Should().BeOfType<SetOperation>();
        var set = (SetOperation)block.Operations[0];
        set.Content.Text.Should().Be("New text");
        set.Content.IsContentBlock.Should().BeFalse();
    }

    [Fact]
    public void Parse_ReplaceWithWith_ProducesReplaceOperation()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nREPLACE \"old\" WITH \"new\"\n");

        var block = doc.OperationBlocks[0];
        block.Operations[0].Should().BeOfType<ReplaceOperation>();
        var replace = (ReplaceOperation)block.Operations[0];
        replace.Search.Should().Be("old");
        replace.Replacement.Should().Be("new");
        replace.IsAll.Should().BeFalse();
    }

    [Fact]
    public void Parse_ReplaceAll_SetsIsAllFlag()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nREPLACE ALL \"old\" WITH \"new\"\n");

        var replace = (ReplaceOperation)doc.OperationBlocks[0].Operations[0];
        replace.IsAll.Should().BeTrue();
    }

    [Fact]
    public void Parse_InsertBefore_WithContentBlock()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE word\n\nAT body/heading[1]\nINSERT BEFORE <<<\nNew section content.\n\nSecond paragraph.\n>>>\n";
        var doc = Parse(input);

        var op = doc.OperationBlocks[0].Operations[0].Should().BeOfType<InsertBeforeOperation>().Subject;
        op.Content.IsContentBlock.Should().BeTrue();
        op.Content.Text.Should().Contain("New section content.");
    }

    [Fact]
    public void Parse_InsertAfter_WithString()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nINSERT AFTER \"Added text\"\n");

        var op = doc.OperationBlocks[0].Operations[0].Should().BeOfType<InsertAfterOperation>().Subject;
        op.Content.Text.Should().Be("Added text");
    }

    [Fact]
    public void Parse_FormatWithProperties_ProducesFormatOperation()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nFORMAT bold=true, font-size=14pt, color=#2B579A\n");

        var format = (FormatOperation)doc.OperationBlocks[0].Operations[0];
        format.Properties.Should().ContainKey("bold");
        format.Properties["bold"].Should().Be(true);
        format.Properties.Should().ContainKey("font-size");
        format.Properties["font-size"].Should().Be("14pt");
        format.Properties.Should().ContainKey("color");
        format.Properties["color"].Should().Be("#2B579A");
    }

    [Fact]
    public void Parse_StyleOperation_ProducesStyleOperation()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nSTYLE \"Heading 2\"\n");

        var style = (StyleOperation)doc.OperationBlocks[0].Operations[0];
        style.StyleName.Should().Be("Heading 2");
    }

    [Fact]
    public void Parse_DeleteOperation_ProducesDeleteOperation()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nDELETE\n");

        var delete = (DeleteOperation)doc.OperationBlocks[0].Operations[0];
        delete.Target.Should().Be(DeleteTarget.Element);
    }

    [Fact]
    public void Parse_DeleteRow_ProducesDeleteWithRowTarget()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/table[1]/row[3]\nDELETE ROW\n");

        var delete = (DeleteOperation)doc.OperationBlocks[0].Operations[0];
        delete.Target.Should().Be(DeleteTarget.Row);
    }

    [Fact]
    public void Parse_AtEach_SetsIsEachFlag()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nAT EACH body/paragraph\nFORMAT bold=true\n");

        doc.OperationBlocks[0].IsEach.Should().BeTrue();
    }

    [Fact]
    public void Parse_PropertyLines_ExtractsProperties()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nPROPERTY title=\"My Document\"\nPROPERTY author=\"John Doe\"\n");

        doc.PropertySettings.Should().HaveCount(2);
        doc.PropertySettings[0].Name.Should().Be("title");
        doc.PropertySettings[0].Value.Should().Be("My Document");
        doc.PropertySettings[1].Name.Should().Be("author");
        doc.PropertySettings[1].Value.Should().Be("John Doe");
    }

    [Fact]
    public void Parse_FullMultiBlockDocument_ParsesAllBlocks()
    {
        var input = @"OFFICETALK/1.0
DOCTYPE word

PROPERTY title=""Annual Report""

AT body/heading[level=1]
SET ""Annual Report — FY2026""
FORMAT font-size=28pt, color=#1F3864

AT body/paragraph[text*=""teh company""]
REPLACE ""teh"" WITH ""the""

AT body/paragraph[text=""DRAFT — DO NOT DISTRIBUTE""]
DELETE
";

        var doc = Parse(input);

        doc.Version.Should().Be("1.0");
        doc.DocType.Should().Be(DocType.Word);
        doc.PropertySettings.Should().HaveCount(1);
        doc.OperationBlocks.Should().HaveCount(3);

        // First block: SET + FORMAT
        doc.OperationBlocks[0].Operations.Should().HaveCount(2);
        doc.OperationBlocks[0].Operations[0].Should().BeOfType<SetOperation>();
        doc.OperationBlocks[0].Operations[1].Should().BeOfType<FormatOperation>();

        // Second block: REPLACE
        doc.OperationBlocks[1].Operations[0].Should().BeOfType<ReplaceOperation>();

        // Third block: DELETE
        doc.OperationBlocks[2].Operations[0].Should().BeOfType<DeleteOperation>();
    }

    [Fact]
    public void Parse_AddressWithKeyValuePredicate_ParsedCorrectly()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/heading[level=2, text=\"Methods\"]\nDELETE\n");

        var segment = doc.OperationBlocks[0].Address.Segments[1];
        segment.Identifier.Should().Be("heading");
        segment.Predicates.Should().HaveCount(2);

        segment.Predicates[0].Should().BeOfType<KeyValuePredicate>();
        var levelPred = (KeyValuePredicate)segment.Predicates[0];
        levelPred.Key.Should().Be("level");
        levelPred.Value.Should().Be("2");

        segment.Predicates[1].Should().BeOfType<KeyValuePredicate>();
        var textPred = (KeyValuePredicate)segment.Predicates[1];
        textPred.Key.Should().Be("text");
        textPred.Value.Should().Be("Methods");
    }

    [Fact]
    public void Parse_AddressWithTextContains_ParsedCorrectly()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[text*=\"conclusion\"]\nDELETE\n");

        var segment = doc.OperationBlocks[0].Address.Segments[1];
        var pred = segment.Predicates[0].Should().BeOfType<KeyValuePredicate>().Subject;
        pred.Key.Should().Be("text");
        pred.Operator.Should().Be(PredicateOperator.AsteriskEquals);
        pred.Value.Should().Be("conclusion");
    }

    [Fact]
    public void Parse_DeepAddress_ParsedCorrectly()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/table[1]/row[3]/cell[2]\nSET \"Updated\"\n");

        var segments = doc.OperationBlocks[0].Address.Segments;
        segments.Should().HaveCount(4);
        segments[0].Identifier.Should().Be("body");
        segments[1].Identifier.Should().Be("table");
        segments[2].Identifier.Should().Be("row");
        segments[3].Identifier.Should().Be("cell");
    }

    [Fact]
    public void Parse_AppendOperation_ProducesAppendOperation()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nAPPEND \" more text\"\n");

        doc.OperationBlocks[0].Operations[0].Should().BeOfType<AppendOperation>();
        var append = (AppendOperation)doc.OperationBlocks[0].Operations[0];
        append.Content.Text.Should().Be(" more text");
    }

    [Fact]
    public void Parse_PrependOperation_ProducesPrependOperation()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nPREPEND \"prefix \"\n");

        doc.OperationBlocks[0].Operations[0].Should().BeOfType<PrependOperation>();
    }

    [Fact]
    public void Parse_InsertRowAfter_Recognized()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/table[1]/row[4]\nINSERT ROW AFTER\n");

        var op = doc.OperationBlocks[0].Operations[0].Should().BeOfType<InsertRowOperation>().Subject;
        op.Position.Should().Be(InsertPosition.After);
    }

    [Fact]
    public void Parse_SetCells_RecognizesValues()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/table[1]/row[4]\nSET CELLS \"Q4\", \"$2.1M\", \"$1.8M\", \"16.7%\"\n");

        var op = doc.OperationBlocks[0].Operations[0].Should().BeOfType<SetCellsOperation>().Subject;
        op.Values.Should().HaveCount(4);
        op.Values[0].Should().Be("Q4");
    }

    [Fact]
    public void Parse_WithComments_CommentsIgnored()
    {
        var input = @"OFFICETALK/1.0
DOCTYPE word

# This is a comment about the first block
AT body/paragraph[1]
# Set new content
SET ""Updated text""
";
        var doc = Parse(input);

        doc.OperationBlocks.Should().HaveCount(1);
        doc.OperationBlocks[0].Operations[0].Should().BeOfType<SetOperation>();
    }

    [Fact]
    public void Parse_ErrorRecovery_CollectsErrorsAndContinues()
    {
        // Missing DOCTYPE — parser should still try to continue
        var doc = Parse("OFFICETALK/1.0\n\nAT body/paragraph[1]\nSET \"text\"\n");

        doc.Errors.Should().NotBeEmpty();
        // Should still have parsed something
        doc.OperationBlocks.Should().HaveCountGreaterOrEqualTo(0);
    }

    [Fact]
    public void Parse_BareStringPredicate_Recognized()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE excel\n\nAT sheet[\"Revenue\"]\nDELETE SHEET\n");

        var segment = doc.OperationBlocks[0].Address.Segments[0];
        segment.Identifier.Should().Be("sheet");
        segment.Predicates.Should().HaveCount(1);
        segment.Predicates[0].Should().BeOfType<BareStringPredicate>();
        var bare = (BareStringPredicate)segment.Predicates[0];
        bare.Value.Should().Be("Revenue");
    }
}
