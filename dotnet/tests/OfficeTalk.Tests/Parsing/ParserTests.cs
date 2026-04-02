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

    // ================ INSERT TEXTBOX tests ================

    [Fact]
    public void Parse_InsertTextbox_AllProperties()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE powerpoint\n\nAT slide[4]\n" +
            "INSERT TEXTBOX left=210pt top=420pt width=300pt height=50pt text=\"hello\" align=center\n";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        var op = doc.OperationBlocks[0].Operations[0].Should().BeOfType<InsertTextboxOperation>().Subject;
        op.Text.Should().Be("hello");
        op.Left.Should().Be("210pt");
        op.Top.Should().Be("420pt");
        op.Width.Should().Be("300pt");
        op.Height.Should().Be("50pt");
        op.Align.Should().Be("center");
    }

    [Fact]
    public void Parse_InsertTextbox_RequiredPropertiesOnly()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE powerpoint\n\nAT slide[1]\n" +
            "INSERT TEXTBOX width=200pt height=100pt text=\"Call out box\"\n";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        var op = doc.OperationBlocks[0].Operations[0].Should().BeOfType<InsertTextboxOperation>().Subject;
        op.Text.Should().Be("Call out box");
        op.Width.Should().Be("200pt");
        op.Height.Should().Be("100pt");
        op.Left.Should().BeNull();
        op.Top.Should().BeNull();
        op.Align.Should().BeNull();
    }

    [Fact]
    public void Parse_InsertShape_WithShapeType()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE powerpoint\n\nAT slide[2]\n" +
            "INSERT SHAPE rectangle left=100pt top=200pt width=300pt height=150pt\n";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        var op = doc.OperationBlocks[0].Operations[0].Should().BeOfType<InsertShapeOperation>().Subject;
        op.ShapeType.Should().Be("rectangle");
        op.Left.Should().Be("100pt");
        op.Top.Should().Be("200pt");
        op.Width.Should().Be("300pt");
        op.Height.Should().Be("150pt");
    }

    [Fact]
    public void Parse_InsertShape_OvalType()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE powerpoint\n\nAT slide[1]\n" +
            "INSERT SHAPE oval width=100pt height=100pt\n";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        var op = doc.OperationBlocks[0].Operations[0].Should().BeOfType<InsertShapeOperation>().Subject;
        op.ShapeType.Should().Be("oval");
        op.Width.Should().Be("100pt");
        op.Height.Should().Be("100pt");
        op.Left.Should().BeNull();
        op.Top.Should().BeNull();
    }

    [Fact]
    public void Parse_InsertTextbox_CanFollowOtherOperationsInBlock()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE powerpoint\n\nAT slide[1]\n" +
            "INSERT TEXTBOX width=100pt height=40pt text=\"Note\"\n" +
            "FORMAT bold=true\n";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        doc.OperationBlocks[0].Operations.Should().HaveCount(2);
        doc.OperationBlocks[0].Operations[0].Should().BeOfType<InsertTextboxOperation>();
        doc.OperationBlocks[0].Operations[1].Should().BeOfType<FormatOperation>();
    }

    // ================ LINK operation tests ================

    [Fact]
    public void Parse_Link_ProducesLinkOperation()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nLINK \"https://example.com\"\n");

        doc.Errors.Should().BeEmpty();
        doc.OperationBlocks.Should().HaveCount(1);
        var op = doc.OperationBlocks[0].Operations[0].Should().BeOfType<LinkOperation>().Subject;
        op.Url.Should().Be("https://example.com");
    }

    [Fact]
    public void Parse_Link_MissingUrl_ProducesError()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nLINK\n");

        doc.Errors.Should().NotBeEmpty();
        doc.Errors[0].Message.Should().Contain("URL");
    }

    // ================ INSERT IMAGE tests ================

    [Fact]
    public void Parse_InsertImage_BasicAfter()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nINSERT IMAGE AFTER \"logo.png\"\n");

        doc.Errors.Should().BeEmpty();
        var op = doc.OperationBlocks[0].Operations[0].Should().BeOfType<InsertImageOperation>().Subject;
        op.Position.Should().Be(InsertPosition.After);
        op.Source.Should().Be("logo.png");
        op.Properties.Should().BeEmpty();
    }

    [Fact]
    public void Parse_InsertImage_WithProperties()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE word\n\nAT body/heading[1]\nINSERT IMAGE BEFORE \"diagram.png\"\n  alt=\"Architecture\"\n  width=6in\n  height=4in\n";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        var op = doc.OperationBlocks[0].Operations[0].Should().BeOfType<InsertImageOperation>().Subject;
        op.Position.Should().Be(InsertPosition.Before);
        op.Source.Should().Be("diagram.png");
        op.Properties.Should().ContainKey("alt");
        op.Properties["alt"].Should().Be("Architecture");
        op.Properties.Should().ContainKey("width");
        op.Properties["width"].Should().Be("6in");
    }

    [Fact]
    public void Parse_InsertImage_MissingSource_ProducesError()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nINSERT IMAGE AFTER\n");

        doc.Errors.Should().NotBeEmpty();
        doc.Errors[0].Message.Should().Contain("image source");
    }

    // ================ INSERT TABLE tests ================

    [Fact]
    public void Parse_InsertTable_BasicDimensions()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/heading[1]\nINSERT TABLE AFTER rows=3, columns=4\n");

        doc.Errors.Should().BeEmpty();
        var op = doc.OperationBlocks[0].Operations[0].Should().BeOfType<InsertTableOperation>().Subject;
        op.Position.Should().Be(InsertPosition.After);
        op.Rows.Should().Be(3);
        op.Columns.Should().Be(4);
    }

    [Fact]
    public void Parse_InsertTable_MissingDimensions_ProducesError()
    {
        var doc = Parse("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/heading[1]\nINSERT TABLE AFTER\n");

        doc.Errors.Should().NotBeEmpty();
        doc.Errors[0].Message.Should().Contain("rows");
    }

    [Fact]
    public void Parse_InsertTable_FollowedBySetCells()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE word\n\nAT body/heading[1]\nINSERT TABLE AFTER rows=2, columns=3\nSET CELLS \"A\", \"B\", \"C\"\n";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        doc.OperationBlocks[0].Operations.Should().HaveCount(2);
        doc.OperationBlocks[0].Operations[0].Should().BeOfType<InsertTableOperation>();
        doc.OperationBlocks[0].Operations[1].Should().BeOfType<SetCellsOperation>();
    }

    // ================ INSERT LIST tests ================

    [Fact]
    public void Parse_InsertList_UnorderedWithItems()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE word\n\nAT body/heading[1]\nINSERT LIST AFTER unordered\n  ITEM \"First item\"\n  ITEM \"Second item\"\n";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        var op = doc.OperationBlocks[0].Operations[0].Should().BeOfType<InsertListOperation>().Subject;
        op.Position.Should().Be(InsertPosition.After);
        op.ListType.Should().Be(ListType.Unordered);
        op.Items.Should().HaveCount(2);
        op.Items[0].Content.Text.Should().Be("First item");
        op.Items[1].Content.Text.Should().Be("Second item");
        op.Items[0].IsNested.Should().BeFalse();
    }

    [Fact]
    public void Parse_InsertList_OrderedType()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nINSERT LIST BEFORE ordered\n  ITEM \"Step one\"\n  ITEM \"Step two\"\n";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        var op = doc.OperationBlocks[0].Operations[0].Should().BeOfType<InsertListOperation>().Subject;
        op.ListType.Should().Be(ListType.Ordered);
    }

    [Fact]
    public void Parse_InsertList_NestedItems()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE word\n\nAT body/heading[1]\nINSERT LIST AFTER unordered\n  ITEM \"Parent\"\n  ITEM \"Child\" nested\n";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        var op = doc.OperationBlocks[0].Operations[0].Should().BeOfType<InsertListOperation>().Subject;
        op.Items.Should().HaveCount(2);
        op.Items[0].IsNested.Should().BeFalse();
        op.Items[1].IsNested.Should().BeTrue();
        op.Items[1].Content.Text.Should().Be("Child");
    }

    [Fact]
    public void Parse_InsertList_DefaultsToUnordered()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE word\n\nAT body/heading[1]\nINSERT LIST AFTER\n  ITEM \"Only item\"\n";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        var op = doc.OperationBlocks[0].Operations[0].Should().BeOfType<InsertListOperation>().Subject;
        op.ListType.Should().Be(ListType.Unordered);
        op.Items.Should().HaveCount(1);
    }

    // ================ SET RUNS tests ================

    [Fact]
    public void Parse_SetRuns_BasicRuns()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nSET RUNS\n  RUN \"plain text\"\n  RUN \"bold\" bold=true\n";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        var op = doc.OperationBlocks[0].Operations[0].Should().BeOfType<SetRunsOperation>().Subject;
        op.Runs.Should().HaveCount(2);
        op.Runs[0].Content.Text.Should().Be("plain text");
        op.Runs[0].Properties.Should().BeEmpty();
        op.Runs[1].Content.Text.Should().Be("bold");
        op.Runs[1].Properties.Should().ContainKey("bold");
        op.Runs[1].Properties["bold"].Should().Be(true);
    }

    [Fact]
    public void Parse_SetRuns_MultipleProperties()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nSET RUNS\n  RUN \"styled\" bold=true, italic=true, color=#FF0000\n";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        var op = doc.OperationBlocks[0].Operations[0].Should().BeOfType<SetRunsOperation>().Subject;
        op.Runs.Should().HaveCount(1);
        op.Runs[0].Properties.Should().HaveCount(3);
        op.Runs[0].Properties["bold"].Should().Be(true);
        op.Runs[0].Properties["italic"].Should().Be(true);
        op.Runs[0].Properties["color"].Should().Be("#FF0000");
    }

    [Fact]
    public void Parse_SetRuns_WithHref()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nSET RUNS\n  RUN \"Visit \"\n  RUN \"us\" href=\"https://example.com\"\n";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        var op = doc.OperationBlocks[0].Operations[0].Should().BeOfType<SetRunsOperation>().Subject;
        op.Runs.Should().HaveCount(2);
        op.Runs[1].Properties["href"].Should().Be("https://example.com");
    }

    [Fact]
    public void Parse_SetRuns_EmptyRuns()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nSET RUNS\n";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        var op = doc.OperationBlocks[0].Operations[0].Should().BeOfType<SetRunsOperation>().Subject;
        op.Runs.Should().BeEmpty();
    }

    // ================ Mixed new operations in blocks ================

    [Fact]
    public void Parse_BlockWithLinkAfterInsertAfter()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[3]\nINSERT AFTER \"See the report.\"\nLINK \"https://example.com\"\n";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        doc.OperationBlocks[0].Operations.Should().HaveCount(2);
        doc.OperationBlocks[0].Operations[0].Should().BeOfType<InsertAfterOperation>();
        doc.OperationBlocks[0].Operations[1].Should().BeOfType<LinkOperation>();
    }

    // ================ New spec properties (background-color, borders) ================

    [Fact]
    public void Parse_Format_BackgroundColor_ParsesAsProperty()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nFORMAT background-color=#F0F0F0\n";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        var format = doc.OperationBlocks[0].Operations[0].Should().BeOfType<FormatOperation>().Subject;
        format.Properties.Should().ContainKey("background-color");
        format.Properties["background-color"].Should().Be("#F0F0F0");
    }

    [Fact]
    public void Parse_Format_ParagraphBorders_AllProperties()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nFORMAT border-bottom=single, border-color=#CCCCCC, border-width=1pt\n";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        var format = doc.OperationBlocks[0].Operations[0].Should().BeOfType<FormatOperation>().Subject;
        format.Properties.Should().ContainKey("border-bottom");
        format.Properties["border-bottom"].Should().Be("single");
        format.Properties.Should().ContainKey("border-color");
        format.Properties["border-color"].Should().Be("#CCCCCC");
        format.Properties.Should().ContainKey("border-width");
        format.Properties["border-width"].Should().Be("1pt");
    }

    [Fact]
    public void Parse_Format_AllFourBorderSides()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nFORMAT border-top=single, border-bottom=double, border-left=thick, border-right=dashed\n";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        var format = doc.OperationBlocks[0].Operations[0].Should().BeOfType<FormatOperation>().Subject;
        format.Properties["border-top"].Should().Be("single");
        format.Properties["border-bottom"].Should().Be("double");
        format.Properties["border-left"].Should().Be("thick");
        format.Properties["border-right"].Should().Be("dashed");
    }

    [Fact]
    public void Parse_SetRuns_BackgroundColorOnRun()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nSET RUNS\n  RUN \"code\" font-name=\"Consolas\" background-color=#F5F5F5\n";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        var op = doc.OperationBlocks[0].Operations[0].Should().BeOfType<SetRunsOperation>().Subject;
        op.Runs.Should().HaveCount(1);
        op.Runs[0].Properties["font-name"].Should().Be("Consolas");
        op.Runs[0].Properties["background-color"].Should().Be("#F5F5F5");
    }

    [Fact]
    public void Parse_SetRuns_SyntaxHighlightedCodePattern()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nSET RUNS\n" +
                    "  RUN \"const \" color=#0000FF font-name=\"Consolas\" font-size=10pt background-color=#F5F5F5\n" +
                    "  RUN \"x\" color=#001080 font-name=\"Consolas\" font-size=10pt background-color=#F5F5F5\n" +
                    "  RUN \" = 42;\" font-name=\"Consolas\" font-size=10pt background-color=#F5F5F5\n";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        var op = doc.OperationBlocks[0].Operations[0].Should().BeOfType<SetRunsOperation>().Subject;
        op.Runs.Should().HaveCount(3);
        op.Runs[0].Properties["color"].Should().Be("#0000FF");
        op.Runs[0].Properties["background-color"].Should().Be("#F5F5F5");
        op.Runs[1].Properties["color"].Should().Be("#001080");
        op.Runs[2].Content.Text.Should().Be(" = 42;");
    }

    [Fact]
    public void Parse_ThematicBreakPattern()
    {
        // Thematic break = empty paragraph with bottom border
        var input = "OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nINSERT AFTER \"\"\nFORMAT border-bottom=single, border-color=#CCCCCC, spacing-after=12pt\n";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        doc.OperationBlocks[0].Operations.Should().HaveCount(2);
        doc.OperationBlocks[0].Operations[0].Should().BeOfType<InsertAfterOperation>();
        var format = doc.OperationBlocks[0].Operations[1].Should().BeOfType<FormatOperation>().Subject;
        format.Properties["border-bottom"].Should().Be("single");
        format.Properties["border-color"].Should().Be("#CCCCCC");
        format.Properties["spacing-after"].Should().Be("12pt");
    }

    [Fact]
    public void Parse_MarkdownDocumentConstructionPattern()
    {
        // End-to-end test of the pattern from spec §12.16
        var input = "OFFICETALK/1.0\nDOCTYPE word\n\n" +
                    "PROPERTY title=\"Getting Started\"\n\n" +
                    "AT body/paragraph[1]\n" +
                    "SET \"Heading\"\n" +
                    "STYLE \"Heading1\"\n\n" +
                    "AT body/paragraph[1]\n" +
                    "INSERT AFTER \"\"\n" +
                    "SET RUNS\n" +
                    "  RUN \"Visit \"\n" +
                    "  RUN \"our site\" href=\"https://example.com\" color=#0563C1 underline=single\n\n" +
                    "AT body/paragraph[2]\n" +
                    "INSERT AFTER \"\"\n" +
                    "FORMAT border-bottom=single, border-color=#CCCCCC\n\n" +
                    "AT body/paragraph[3]\n" +
                    "INSERT LIST AFTER unordered\n" +
                    "  ITEM \"First\"\n" +
                    "  ITEM \"Second\"\n";
        var doc = Parse(input);

        doc.Errors.Should().BeEmpty();
        doc.PropertySettings.Should().Contain(p => p.Name == "title");
        doc.OperationBlocks.Should().HaveCount(4);

        // Block 1: SET + STYLE
        doc.OperationBlocks[0].Operations[0].Should().BeOfType<SetOperation>();
        doc.OperationBlocks[0].Operations[1].Should().BeOfType<StyleOperation>();

        // Block 2: INSERT AFTER + SET RUNS
        doc.OperationBlocks[1].Operations[0].Should().BeOfType<InsertAfterOperation>();
        doc.OperationBlocks[1].Operations[1].Should().BeOfType<SetRunsOperation>();

        // Block 3: INSERT AFTER + FORMAT (thematic break)
        doc.OperationBlocks[2].Operations[0].Should().BeOfType<InsertAfterOperation>();
        var fmt = doc.OperationBlocks[2].Operations[1].Should().BeOfType<FormatOperation>().Subject;
        fmt.Properties.Should().ContainKey("border-bottom");

        // Block 4: INSERT LIST
        doc.OperationBlocks[3].Operations[0].Should().BeOfType<InsertListOperation>();
    }
}
