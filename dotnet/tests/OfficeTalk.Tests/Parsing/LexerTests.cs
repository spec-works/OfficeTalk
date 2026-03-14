using FluentAssertions;
using OfficeTalk.Parsing;

namespace OfficeTalk.Tests.Parsing;

public class LexerTests
{
    [Fact]
    public void Tokenize_VersionHeader_ProducesVersionToken()
    {
        var lexer = new OfficeTalkLexer("OFFICETALK/1.0\nDOCTYPE word\n");
        var tokens = lexer.Tokenize();

        tokens[0].Type.Should().Be(TokenType.Version);
        tokens[0].Value.Should().Be("OFFICETALK/1.0");
    }

    [Fact]
    public void Tokenize_DocType_ProducesDocTypeTokens()
    {
        var lexer = new OfficeTalkLexer("OFFICETALK/1.0\nDOCTYPE word\n");
        var tokens = lexer.Tokenize();

        var docTypeToken = tokens.First(t => t.Type == TokenType.DocTypeKeyword);
        docTypeToken.Value.Should().Be("DOCTYPE");

        var wordToken = tokens.First(t => t.Type == TokenType.Word);
        wordToken.Value.Should().Be("word");
    }

    [Theory]
    [InlineData("excel", TokenType.Excel)]
    [InlineData("powerpoint", TokenType.PowerPoint)]
    [InlineData("word", TokenType.Word)]
    public void Tokenize_AllDocTypes_RecognizedCorrectly(string docType, TokenType expected)
    {
        var lexer = new OfficeTalkLexer($"OFFICETALK/1.0\nDOCTYPE {docType}\n");
        var tokens = lexer.Tokenize();

        tokens.Should().Contain(t => t.Type == expected);
    }

    [Fact]
    public void Tokenize_SimpleString_ProducesStringToken()
    {
        var lexer = new OfficeTalkLexer("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nSET \"Hello, World!\"\n");
        var tokens = lexer.Tokenize();

        var stringToken = tokens.First(t => t.Type == TokenType.String);
        stringToken.Value.Should().Be("Hello, World!");
    }

    [Fact]
    public void Tokenize_StringWithEscapes_ProcessesEscapeSequences()
    {
        var lexer = new OfficeTalkLexer("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nSET \"She said \\\"hello\\\" to them.\"\n");
        var tokens = lexer.Tokenize();

        var stringToken = tokens.First(t => t.Type == TokenType.String);
        stringToken.Value.Should().Be("She said \"hello\" to them.");
    }

    [Fact]
    public void Tokenize_StringWithNewlineEscape_ProducesNewlineCharacter()
    {
        var lexer = new OfficeTalkLexer("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nSET \"Line one\\nLine two\"\n");
        var tokens = lexer.Tokenize();

        var stringToken = tokens.First(t => t.Type == TokenType.String);
        stringToken.Value.Should().Be("Line one\nLine two");
    }

    [Fact]
    public void Tokenize_Address_ProducesCorrectSegments()
    {
        var lexer = new OfficeTalkLexer("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[3]\nDELETE\n");
        var tokens = lexer.Tokenize();

        tokens.Should().Contain(t => t.Type == TokenType.AT);
        tokens.Should().Contain(t => t.Type == TokenType.Identifier && t.Value == "body");
        tokens.Should().Contain(t => t.Type == TokenType.Slash);
        tokens.Should().Contain(t => t.Type == TokenType.Identifier && t.Value == "paragraph");
        tokens.Should().Contain(t => t.Type == TokenType.LeftBracket);
        tokens.Should().Contain(t => t.Type == TokenType.Number && t.Value == "3");
        tokens.Should().Contain(t => t.Type == TokenType.RightBracket);
    }

    [Fact]
    public void Tokenize_ContentBlock_CapturedAsOneToken()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nINSERT BEFORE <<<\nFirst paragraph.\n\nSecond paragraph.\n>>>\n";
        var lexer = new OfficeTalkLexer(input);
        var tokens = lexer.Tokenize();

        var contentToken = tokens.First(t => t.Type == TokenType.ContentBlock);
        contentToken.Value.Should().Contain("First paragraph.");
        contentToken.Value.Should().Contain("Second paragraph.");
    }

    [Fact]
    public void Tokenize_Comment_CapturedAsCommentToken()
    {
        var lexer = new OfficeTalkLexer("OFFICETALK/1.0\nDOCTYPE word\n\n# This is a comment\nAT body/paragraph[1]\nDELETE\n");
        var tokens = lexer.Tokenize();

        tokens.Should().Contain(t => t.Type == TokenType.Comment && t.Value.Contains("This is a comment"));
    }

    [Fact]
    public void Tokenize_Number_ProducesNumberToken()
    {
        var lexer = new OfficeTalkLexer("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/heading[level=2]\nDELETE\n");
        var tokens = lexer.Tokenize();

        tokens.Should().Contain(t => t.Type == TokenType.Number && t.Value == "2");
    }

    [Fact]
    public void Tokenize_Boolean_ProducesBooleanToken()
    {
        var lexer = new OfficeTalkLexer("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nFORMAT bold=true\n");
        var tokens = lexer.Tokenize();

        tokens.Should().Contain(t => t.Type == TokenType.Boolean && t.Value == "true");
    }

    [Fact]
    public void Tokenize_Color_ProducesColorToken()
    {
        var lexer = new OfficeTalkLexer("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nFORMAT color=#2B579A\n");
        var tokens = lexer.Tokenize();

        tokens.Should().Contain(t => t.Type == TokenType.Color && t.Value == "#2B579A");
    }

    [Fact]
    public void Tokenize_Length_ProducesLengthToken()
    {
        var lexer = new OfficeTalkLexer("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nFORMAT font-size=14pt\n");
        var tokens = lexer.Tokenize();

        tokens.Should().Contain(t => t.Type == TokenType.Length && t.Value == "14pt");
    }

    [Theory]
    [InlineData("12pt")]
    [InlineData("1.5in")]
    [InlineData("2.54cm")]
    [InlineData("50%")]
    public void Tokenize_VariousLengths_RecognizedCorrectly(string lengthValue)
    {
        var lexer = new OfficeTalkLexer($"OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nFORMAT font-size={lengthValue}\n");
        var tokens = lexer.Tokenize();

        tokens.Should().Contain(t => t.Type == TokenType.Length && t.Value == lengthValue);
    }

    [Fact]
    public void Tokenize_ComparisonOperators_RecognizedCorrectly()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[text*=\"hello\"]\nDELETE\n";
        var lexer = new OfficeTalkLexer(input);
        var tokens = lexer.Tokenize();

        tokens.Should().Contain(t => t.Type == TokenType.AsteriskEquals);
    }

    [Theory]
    [InlineData("~=", TokenType.TildeEquals)]
    [InlineData("^=", TokenType.CaretEquals)]
    [InlineData("$=", TokenType.DollarEquals)]
    [InlineData("*=", TokenType.AsteriskEquals)]
    public void Tokenize_AllComparisonOperators_Recognized(string op, TokenType expected)
    {
        var input = $"OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[text{op}\"test\"]\nDELETE\n";
        var lexer = new OfficeTalkLexer(input);
        var tokens = lexer.Tokenize();

        tokens.Should().Contain(t => t.Type == expected);
    }

    [Fact]
    public void Tokenize_Keywords_RecognizedCorrectly()
    {
        var input = "OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nREPLACE ALL \"old\" WITH \"new\"\n";
        var lexer = new OfficeTalkLexer(input);
        var tokens = lexer.Tokenize();

        tokens.Should().Contain(t => t.Type == TokenType.REPLACE);
        tokens.Should().Contain(t => t.Type == TokenType.ALL);
        tokens.Should().Contain(t => t.Type == TokenType.WITH);
    }

    [Fact]
    public void Tokenize_TracksLineAndColumn()
    {
        var lexer = new OfficeTalkLexer("OFFICETALK/1.0\nDOCTYPE word\n");
        var tokens = lexer.Tokenize();

        tokens[0].Line.Should().Be(1); // version on line 1
        var docTypeToken = tokens.First(t => t.Type == TokenType.DocTypeKeyword);
        docTypeToken.Line.Should().Be(2);
    }

    [Fact]
    public void Tokenize_EmptyInput_ProducesVersionAndEof()
    {
        var lexer = new OfficeTalkLexer("");
        var tokens = lexer.Tokenize();

        // Should have at least version (empty) and EOF
        tokens.Should().Contain(t => t.Type == TokenType.EOF);
    }

    [Fact]
    public void Tokenize_ColorWithAlpha_RecognizedCorrectly()
    {
        var lexer = new OfficeTalkLexer("OFFICETALK/1.0\nDOCTYPE word\n\nAT body/paragraph[1]\nFORMAT color=#FF000080\n");
        var tokens = lexer.Tokenize();

        tokens.Should().Contain(t => t.Type == TokenType.Color && t.Value == "#FF000080");
    }
}
