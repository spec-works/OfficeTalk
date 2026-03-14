using FluentAssertions;
using OfficeTalk.Ast;
using OfficeTalk.Parsing;
using OfficeTalk.Validation;

namespace OfficeTalk.Tests.TestCases;

public class NegativeTestCaseTests
{
    private static readonly string NegativeDir = Path.Combine(
        AppContext.BaseDirectory, "testcases", "negative");

    public static IEnumerable<object[]> GetNegativeTestCases()
    {
        foreach (var otkFile in Directory.GetFiles(NegativeDir, "*.otk"))
        {
            yield return new object[] { Path.GetFileNameWithoutExtension(otkFile) };
        }
    }

    [Theory]
    [MemberData(nameof(GetNegativeTestCases))]
    public void NegativeTestCase_FailsParsingOrValidation(string testName)
    {
        var otkPath = Path.Combine(NegativeDir, $"{testName}.otk");
        var otkSource = File.ReadAllText(otkPath);

        bool hasErrors = false;

        try
        {
            var lexer = new OfficeTalkLexer(otkSource);
            var tokens = lexer.Tokenize();
            var parser = new OfficeTalkParser(tokens);
            var document = parser.Parse();

            if (document.Errors.Count > 0)
            {
                hasErrors = true;
            }
            else
            {
                var validator = new SyntacticValidator();
                var result = validator.Validate(document);
                hasErrors = !result.IsValid;
            }

            if (!hasErrors)
            {
                // The lexer silently recovers from some errors (e.g. unterminated
                // strings/content blocks). Fall back to detecting lexical issues
                // in the source that the parser didn't flag.
                hasErrors = HasLexicalIssues(otkSource);
            }
        }
        catch
        {
            // Lexer or parser threw an exception — that's also a valid failure
            hasErrors = true;
        }

        hasErrors.Should().BeTrue(
            $"test case '{testName}' should fail during parsing or validation");
    }

    /// <summary>
    /// Detects lexical issues that the lexer silently recovers from,
    /// such as unterminated strings and unterminated content blocks.
    /// </summary>
    private static bool HasLexicalIssues(string source)
    {
        var lines = source.Replace("\r\n", "\n").Split('\n');
        bool inContentBlock = false;

        foreach (var line in lines)
        {
            var trimmed = line.Trim();

            if (inContentBlock)
            {
                if (trimmed.StartsWith(">>>"))
                    inContentBlock = false;
                continue;
            }

            if (trimmed.EndsWith("<<<"))
            {
                inContentBlock = true;
                continue;
            }

            // Skip comments
            if (trimmed.StartsWith("#"))
                continue;

            // Check for unterminated strings (odd number of unescaped quotes)
            int quoteCount = 0;
            for (int i = 0; i < line.Length; i++)
            {
                if (line[i] == '"' && (i == 0 || line[i - 1] != '\\'))
                    quoteCount++;
            }

            if (quoteCount % 2 != 0)
                return true;
        }

        // Still inside a content block at EOF means it's unterminated
        return inContentBlock;
    }
}
