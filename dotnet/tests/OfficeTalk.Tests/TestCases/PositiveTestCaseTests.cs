using System.Text.Json;
using System.Text.RegularExpressions;
using FluentAssertions;
using OfficeTalk.Ast;
using OfficeTalk.Parsing;

namespace OfficeTalk.Tests.TestCases;

public class PositiveTestCaseTests
{
    private static readonly string TestCasesDir = Path.Combine(
        AppContext.BaseDirectory, "testcases");

    public static IEnumerable<object[]> GetPositiveTestCases()
    {
        foreach (var otkFile in Directory.GetFiles(TestCasesDir, "*.otk"))
        {
            var jsonFile = Path.ChangeExtension(otkFile, ".json");
            if (File.Exists(jsonFile))
            {
                yield return new object[] { Path.GetFileNameWithoutExtension(otkFile) };
            }
        }
    }

    [Theory]
    [MemberData(nameof(GetPositiveTestCases))]
    public void PositiveTestCase_ParsesCorrectly(string testName)
    {
        var otkPath = Path.Combine(TestCasesDir, $"{testName}.otk");
        var jsonPath = Path.Combine(TestCasesDir, $"{testName}.json");

        var otkSource = File.ReadAllText(otkPath);
        var jsonSource = File.ReadAllText(jsonPath);

        // Parse the OTK file
        var lexer = new OfficeTalkLexer(otkSource);
        var tokens = lexer.Tokenize();
        var parser = new OfficeTalkParser(tokens);
        var document = parser.Parse();

        // Deserialize the expected JSON
        using var jsonDoc = JsonDocument.Parse(jsonSource);
        var expected = jsonDoc.RootElement;

        // Verify no parse errors
        document.Errors.Should().BeEmpty($"test case '{testName}' should parse without errors");

        // Verify version
        document.Version.Should().Be(expected.GetProperty("version").GetString(),
            $"version should match in test case '{testName}'");

        // Verify docType
        var expectedDocType = expected.GetProperty("docType").GetString();
        document.DocType.ToString().ToLowerInvariant().Should().Be(expectedDocType,
            $"docType should match in test case '{testName}'");

        // Verify operation blocks
        var expectedBlocks = expected.GetProperty("operationBlocks");
        document.OperationBlocks.Should().HaveCount(expectedBlocks.GetArrayLength(),
            $"operation block count should match in test case '{testName}'");

        for (int i = 0; i < expectedBlocks.GetArrayLength(); i++)
        {
            var expectedBlock = expectedBlocks[i];
            var actualBlock = document.OperationBlocks[i];
            var blockContext = $"block {i} in test case '{testName}'";

            // Verify address — normalize bare numeric values in key-value predicates
            // because Address.ToString() always quotes predicate values (e.g. level="1")
            // while fixture JSON uses bare numbers (e.g. level=1).
            var expectedAddress = NormalizeExpectedAddress(
                expectedBlock.GetProperty("address").GetString()!);
            actualBlock.Address.ToString().Should().Be(expectedAddress,
                $"address should match for {blockContext}");

            // Verify each flag
            actualBlock.IsEach.Should().Be(expectedBlock.GetProperty("each").GetBoolean(),
                $"'each' flag should match for {blockContext}");

            // Verify operations
            var expectedOps = expectedBlock.GetProperty("operations");
            actualBlock.Operations.Should().HaveCount(expectedOps.GetArrayLength(),
                $"operation count should match for {blockContext}");

            for (int j = 0; j < expectedOps.GetArrayLength(); j++)
            {
                VerifyOperation(actualBlock.Operations[j], expectedOps[j], testName, i, j);
            }
        }

        // Verify property settings
        var expectedProps = expected.GetProperty("propertySettings");
        document.PropertySettings.Should().HaveCount(expectedProps.GetArrayLength(),
            $"property settings count should match in test case '{testName}'");

        for (int i = 0; i < expectedProps.GetArrayLength(); i++)
        {
            var expectedProp = expectedProps[i];
            var propContext = $"property {i} in test case '{testName}'";
            document.PropertySettings[i].Name.Should().Be(
                expectedProp.GetProperty("name").GetString(), propContext);
            document.PropertySettings[i].Value.Should().Be(
                expectedProp.GetProperty("value").GetString(), propContext);
        }
    }

    private static void VerifyOperation(Operation actual, JsonElement expected,
        string testName, int blockIdx, int opIdx)
    {
        var context = $"block {blockIdx}, operation {opIdx} in test case '{testName}'";
        var expectedType = expected.GetProperty("type").GetString();

        switch (expectedType)
        {
            case "SET":
                actual.Should().BeOfType<SetOperation>(context);
                var setOp = (SetOperation)actual;
                NormalizeNewlines(setOp.Content.Text).Should().Be(
                    NormalizeNewlines(expected.GetProperty("content").GetString()!), context);
                break;

            case "REPLACE":
                actual.Should().BeOfType<ReplaceOperation>(context);
                var replaceOp = (ReplaceOperation)actual;
                replaceOp.Search.Should().Be(expected.GetProperty("search").GetString(), context);
                replaceOp.Replacement.Should().Be(expected.GetProperty("replacement").GetString(), context);
                replaceOp.IsAll.Should().Be(expected.GetProperty("all").GetBoolean(), context);
                break;

            case "DELETE":
                actual.Should().BeOfType<DeleteOperation>(context);
                break;

            case "FORMAT":
                actual.Should().BeOfType<FormatOperation>(context);
                VerifyFormatProperties((FormatOperation)actual, expected.GetProperty("properties"), context);
                break;

            case "STYLE":
                actual.Should().BeOfType<StyleOperation>(context);
                ((StyleOperation)actual).StyleName.Should().Be(
                    expected.GetProperty("styleName").GetString(), context);
                break;

            case "INSERT_BEFORE":
                actual.Should().BeOfType<InsertBeforeOperation>(context);
                NormalizeNewlines(((InsertBeforeOperation)actual).Content.Text).Should().Be(
                    NormalizeNewlines(expected.GetProperty("content").GetString()!), context);
                break;

            case "INSERT_AFTER":
                actual.Should().BeOfType<InsertAfterOperation>(context);
                NormalizeNewlines(((InsertAfterOperation)actual).Content.Text).Should().Be(
                    NormalizeNewlines(expected.GetProperty("content").GetString()!), context);
                break;

            case "APPEND":
                actual.Should().BeOfType<AppendOperation>(context);
                NormalizeNewlines(((AppendOperation)actual).Content.Text).Should().Be(
                    NormalizeNewlines(expected.GetProperty("content").GetString()!), context);
                break;

            case "PREPEND":
                actual.Should().BeOfType<PrependOperation>(context);
                NormalizeNewlines(((PrependOperation)actual).Content.Text).Should().Be(
                    NormalizeNewlines(expected.GetProperty("content").GetString()!), context);
                break;

            case "INSERT_ROW":
                actual.Should().BeOfType<InsertRowOperation>(context);
                var insertRowOp = (InsertRowOperation)actual;
                insertRowOp.Position.ToString().ToUpperInvariant().Should().Be(
                    expected.GetProperty("position").GetString(), context);
                break;

            case "INSERT_COLUMN":
                actual.Should().BeOfType<InsertColumnOperation>(context);
                var insertColOp = (InsertColumnOperation)actual;
                insertColOp.Position.ToString().ToUpperInvariant().Should().Be(
                    expected.GetProperty("position").GetString(), context);
                break;

            case "INSERT_TEXTBOX":
                actual.Should().BeOfType<InsertTextboxOperation>(context);
                var insertTextboxOp = (InsertTextboxOperation)actual;
                insertTextboxOp.Text.Should().Be(expected.GetProperty("text").GetString(), context);
                if (expected.TryGetProperty("left", out var tbLeft))
                    insertTextboxOp.Left.Should().Be(tbLeft.GetString(), context);
                if (expected.TryGetProperty("top", out var tbTop))
                    insertTextboxOp.Top.Should().Be(tbTop.GetString(), context);
                if (expected.TryGetProperty("width", out var tbWidth))
                    insertTextboxOp.Width.Should().Be(tbWidth.GetString(), context);
                if (expected.TryGetProperty("height", out var tbHeight))
                    insertTextboxOp.Height.Should().Be(tbHeight.GetString(), context);
                if (expected.TryGetProperty("align", out var tbAlign))
                    insertTextboxOp.Align.Should().Be(tbAlign.GetString(), context);
                break;

            case "INSERT_SHAPE":
                actual.Should().BeOfType<InsertShapeOperation>(context);
                var insertShapeOp = (InsertShapeOperation)actual;
                insertShapeOp.ShapeType.Should().Be(expected.GetProperty("shapeType").GetString(), context);
                if (expected.TryGetProperty("left", out var shLeft))
                    insertShapeOp.Left.Should().Be(shLeft.GetString(), context);
                if (expected.TryGetProperty("top", out var shTop))
                    insertShapeOp.Top.Should().Be(shTop.GetString(), context);
                if (expected.TryGetProperty("width", out var shWidth))
                    insertShapeOp.Width.Should().Be(shWidth.GetString(), context);
                if (expected.TryGetProperty("height", out var shHeight))
                    insertShapeOp.Height.Should().Be(shHeight.GetString(), context);
                break;

            case "SET_CELLS":
                actual.Should().BeOfType<SetCellsOperation>(context);
                var setCellsOp = (SetCellsOperation)actual;
                var expectedValues = expected.GetProperty("values");
                setCellsOp.Values.Should().HaveCount(expectedValues.GetArrayLength(), context);
                for (int k = 0; k < expectedValues.GetArrayLength(); k++)
                {
                    setCellsOp.Values[k].Should().Be(expectedValues[k].GetString(),
                        $"value {k} in {context}");
                }
                break;

            default:
                Assert.Fail($"Unknown operation type '{expectedType}' in {context}");
                break;
        }
    }

    private static void VerifyFormatProperties(FormatOperation formatOp,
        JsonElement expectedProps, string context)
    {
        foreach (var prop in expectedProps.EnumerateObject())
        {
            formatOp.Properties.Should().ContainKey(prop.Name,
                $"property '{prop.Name}' should exist in {context}");
            var actualValue = formatOp.Properties[prop.Name];

            switch (prop.Value.ValueKind)
            {
                case JsonValueKind.True:
                    actualValue.Should().Be(true, $"property '{prop.Name}' in {context}");
                    break;
                case JsonValueKind.False:
                    actualValue.Should().Be(false, $"property '{prop.Name}' in {context}");
                    break;
                case JsonValueKind.String:
                    actualValue?.ToString().Should().Be(prop.Value.GetString(),
                        $"property '{prop.Name}' in {context}");
                    break;
                case JsonValueKind.Number:
                    actualValue?.ToString().Should().Be(prop.Value.GetRawText(),
                        $"property '{prop.Name}' in {context}");
                    break;
                default:
                    Assert.Fail($"Unexpected JSON value kind {prop.Value.ValueKind} for " +
                        $"property '{prop.Name}' in {context}");
                    break;
            }
        }
    }

    private static string NormalizeNewlines(string s) => s.Replace("\r\n", "\n");

    /// <summary>
    /// Address.ToString() wraps all key-value predicate values in quotes,
    /// but fixture JSON may use bare numbers (e.g. level=2 instead of level="2").
    /// This normalizes the expected address to match ToString() output.
    /// </summary>
    private static string NormalizeExpectedAddress(string address) =>
        Regex.Replace(address, @"(~=|\^=|\$=|\*=|=)(\d+)(?=[,\]])", "$1\"$2\"");
}
