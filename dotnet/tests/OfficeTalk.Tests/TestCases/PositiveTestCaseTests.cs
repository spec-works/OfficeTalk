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

        // Verify inspect blocks (if present in the expected JSON)
        if (expected.TryGetProperty("inspectBlocks", out var expectedInspects))
        {
            document.InspectBlocks.Should().HaveCount(expectedInspects.GetArrayLength(),
                $"inspect block count should match in test case '{testName}'");

            for (int i = 0; i < expectedInspects.GetArrayLength(); i++)
            {
                var expectedBlock = expectedInspects[i];
                var actualBlock = document.InspectBlocks[i];
                var blockContext = $"inspect block {i} in test case '{testName}'";

                // Verify address
                var expectedAddress = NormalizeExpectedAddress(
                    expectedBlock.GetProperty("address").GetString()!);
                actualBlock.Address.ToString().Should().Be(expectedAddress,
                    $"address should match for {blockContext}");

                // Verify depth
                actualBlock.Depth.Should().Be(
                    expectedBlock.GetProperty("depth").GetInt32(),
                    $"depth should match for {blockContext}");

                // Verify include layers
                var expectedInclude = expectedBlock.GetProperty("include");
                var expectedLayers = new List<string>();
                foreach (var layer in expectedInclude.EnumerateArray())
                {
                    expectedLayers.Add(layer.GetString()!);
                }

                actualBlock.Include.Should().HaveCount(expectedLayers.Count,
                    $"include layer count should match for {blockContext}");
                for (int j = 0; j < expectedLayers.Count; j++)
                {
                    actualBlock.Include[j].ToString().ToLowerInvariant().Should().Be(
                        expectedLayers[j],
                        $"include layer {j} should match for {blockContext}");
                }

                // Verify context
                actualBlock.Context.Should().Be(
                    expectedBlock.GetProperty("context").GetInt32(),
                    $"context should match for {blockContext}");
            }
        }
        else
        {
            // If no inspectBlocks in JSON, document should have none
            document.InspectBlocks.Should().BeEmpty(
                $"no inspect blocks expected in test case '{testName}'");
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

            case "LINK":
                actual.Should().BeOfType<LinkOperation>(context);
                ((LinkOperation)actual).Url.Should().Be(
                    expected.GetProperty("url").GetString(), context);
                break;

            case "INSERT_IMAGE":
                actual.Should().BeOfType<InsertImageOperation>(context);
                var imgOp = (InsertImageOperation)actual;
                imgOp.Position.ToString().ToUpperInvariant().Should().Be(
                    expected.GetProperty("position").GetString(), context);
                imgOp.Source.Should().Be(expected.GetProperty("source").GetString(), context);
                if (expected.TryGetProperty("properties", out var imgProps))
                    VerifyFormatProperties2(imgOp.Properties, imgProps, context);
                break;

            case "INSERT_TABLE":
                actual.Should().BeOfType<InsertTableOperation>(context);
                var tableOp = (InsertTableOperation)actual;
                tableOp.Position.ToString().ToUpperInvariant().Should().Be(
                    expected.GetProperty("position").GetString(), context);
                tableOp.Rows.Should().Be(expected.GetProperty("rows").GetInt32(), context);
                tableOp.Columns.Should().Be(expected.GetProperty("columns").GetInt32(), context);
                break;

            case "INSERT_LIST":
                actual.Should().BeOfType<InsertListOperation>(context);
                var listOp = (InsertListOperation)actual;
                listOp.Position.ToString().ToUpperInvariant().Should().Be(
                    expected.GetProperty("position").GetString(), context);
                listOp.ListType.ToString().ToUpperInvariant().Should().Be(
                    expected.GetProperty("listType").GetString()!.ToUpperInvariant(), context);
                var expectedItems = expected.GetProperty("items");
                listOp.Items.Should().HaveCount(expectedItems.GetArrayLength(), context);
                for (int k = 0; k < expectedItems.GetArrayLength(); k++)
                {
                    var item = expectedItems[k];
                    NormalizeNewlines(listOp.Items[k].Content.Text).Should().Be(
                        NormalizeNewlines(item.GetProperty("content").GetString()!),
                        $"item {k} in {context}");
                    listOp.Items[k].IsNested.Should().Be(
                        item.TryGetProperty("nested", out var nested) && nested.GetBoolean(),
                        $"item {k} nested in {context}");
                }
                break;

            case "SET_RUNS":
                actual.Should().BeOfType<SetRunsOperation>(context);
                var runsOp = (SetRunsOperation)actual;
                var expectedRuns = expected.GetProperty("runs");
                runsOp.Runs.Should().HaveCount(expectedRuns.GetArrayLength(), context);
                for (int k = 0; k < expectedRuns.GetArrayLength(); k++)
                {
                    var run = expectedRuns[k];
                    NormalizeNewlines(runsOp.Runs[k].Content.Text).Should().Be(
                        NormalizeNewlines(run.GetProperty("content").GetString()!),
                        $"run {k} in {context}");
                    if (run.TryGetProperty("properties", out var runProps))
                        VerifyFormatProperties2(runsOp.Runs[k].Properties, runProps,
                            $"run {k} in {context}");
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
        VerifyFormatProperties2(formatOp.Properties, expectedProps, context);
    }

    private static void VerifyFormatProperties2(Dictionary<string, object> properties,
        JsonElement expectedProps, string context)
    {
        foreach (var prop in expectedProps.EnumerateObject())
        {
            properties.Should().ContainKey(prop.Name,
                $"property '{prop.Name}' should exist in {context}");
            var actualValue = properties[prop.Name];

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
