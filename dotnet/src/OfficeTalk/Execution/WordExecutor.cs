using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeTalk.Addressing;
using OfficeTalk.Ast;

namespace OfficeTalk.Execution;

/// <summary>
/// Executes OfficeTalk operations against a Word (docx) document.
/// Uses snapshot semantics: all addresses are resolved before any operations are applied.
/// </summary>
public class WordExecutor : IOfficeTalkExecutor
{
    /// <inheritdoc/>
    public void Execute(OfficeTalkDocument document, string targetPath, string? outputPath = null)
    {
        var workingPath = outputPath ?? targetPath;

        if (outputPath != null && outputPath != targetPath)
            File.Copy(targetPath, outputPath, overwrite: true);

        using var wordDoc = WordprocessingDocument.Open(workingPath, true);

        var resolver = new WordAddressResolver(wordDoc);

        // Phase 1: Resolve all addresses (snapshot semantics)
        var resolvedBlocks = new List<(OperationBlock Block, IReadOnlyList<OpenXmlElement> Elements)>();
        foreach (var block in document.OperationBlocks)
        {
            var elements = resolver.Resolve(block.Address);
            resolvedBlocks.Add((block, elements));
        }

        // Phase 2: Apply property settings
        foreach (var prop in document.PropertySettings)
        {
            ApplyProperty(wordDoc, prop);
        }

        // Phase 3: Execute operations
        foreach (var (block, elements) in resolvedBlocks)
        {
            foreach (var element in elements)
            {
                foreach (var operation in block.Operations)
                {
                    ExecuteOperation(element, operation);
                }
            }
        }

        wordDoc.Save();
    }

    private static void ApplyProperty(WordprocessingDocument wordDoc, PropertySetting prop)
    {
        var coreProps = wordDoc.PackageProperties;
        switch (prop.Name.ToLowerInvariant())
        {
            case "title":
                coreProps.Title = prop.Value;
                break;
            case "author":
                coreProps.Creator = prop.Value;
                break;
            case "subject":
                coreProps.Subject = prop.Value;
                break;
            case "description":
                coreProps.Description = prop.Value;
                break;
        }
    }

    private static void ExecuteOperation(OpenXmlElement element, Operation operation)
    {
        switch (operation)
        {
            case SetOperation set:
                ExecuteSet(element, set);
                break;
            case ReplaceOperation replace:
                ExecuteReplace(element, replace);
                break;
            case DeleteOperation:
                element.Remove();
                break;
            case AppendOperation append:
                ExecuteAppend(element, append);
                break;
            case PrependOperation prepend:
                ExecutePrepend(element, prepend);
                break;
            case StyleOperation style:
                ExecuteStyle(element, style);
                break;
            case FormatOperation:
                throw new NotImplementedException("FORMAT execution is not yet implemented.");
            case InsertBeforeOperation:
                throw new NotImplementedException("INSERT BEFORE execution is not yet implemented.");
            case InsertAfterOperation:
                throw new NotImplementedException("INSERT AFTER execution is not yet implemented.");
            case InsertRowOperation:
                throw new NotImplementedException("INSERT ROW execution is not yet implemented.");
            case InsertColumnOperation:
                throw new NotImplementedException("INSERT COLUMN execution is not yet implemented.");
            case MergeCellsOperation:
                throw new NotImplementedException("MERGE CELLS execution is not yet implemented.");
            case SetCellsOperation:
                throw new NotImplementedException("SET CELLS execution is not yet implemented.");
            case InsertSlideOperation:
                throw new NotImplementedException("INSERT SLIDE execution is not yet implemented.");
            case DuplicateSlideOperation:
                throw new NotImplementedException("DUPLICATE SLIDE execution is not yet implemented.");
            case RenameSheetOperation:
                throw new NotImplementedException("RENAME SHEET execution is not yet implemented.");
            case AddSheetOperation:
                throw new NotImplementedException("ADD SHEET execution is not yet implemented.");
        }
    }

    private static void ExecuteSet(OpenXmlElement element, SetOperation set)
    {
        if (element is Paragraph paragraph)
        {
            // Remove all existing runs
            paragraph.RemoveAllChildren<Run>();

            // Add new run with content
            var run = new Run(new Text(set.Content.Text) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve });
            paragraph.AppendChild(run);
        }
    }

    private static void ExecuteReplace(OpenXmlElement element, ReplaceOperation replace)
    {
        if (element is Paragraph paragraph)
        {
            foreach (var run in paragraph.Elements<Run>().ToList())
            {
                foreach (var text in run.Elements<Text>().ToList())
                {
                    if (replace.IsAll)
                    {
                        text.Text = text.Text.Replace(replace.Search, replace.Replacement);
                    }
                    else
                    {
                        var idx = text.Text.IndexOf(replace.Search, StringComparison.Ordinal);
                        if (idx >= 0)
                        {
                            text.Text = string.Concat(
                                text.Text.AsSpan(0, idx),
                                replace.Replacement,
                                text.Text.AsSpan(idx + replace.Search.Length));
                        }
                    }
                }
            }
        }
    }

    private static void ExecuteAppend(OpenXmlElement element, AppendOperation append)
    {
        if (element is Paragraph paragraph)
        {
            var run = new Run(new Text(append.Content.Text) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve });
            paragraph.AppendChild(run);
        }
    }

    private static void ExecutePrepend(OpenXmlElement element, PrependOperation prepend)
    {
        if (element is Paragraph paragraph)
        {
            var run = new Run(new Text(prepend.Content.Text) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve });
            var firstRun = paragraph.Elements<Run>().FirstOrDefault();
            if (firstRun != null)
                paragraph.InsertBefore(run, firstRun);
            else
                paragraph.AppendChild(run);
        }
    }

    private static void ExecuteStyle(OpenXmlElement element, StyleOperation style)
    {
        if (element is Paragraph paragraph)
        {
            var props = paragraph.ParagraphProperties ?? paragraph.PrependChild(new ParagraphProperties());
            props.ParagraphStyleId = new ParagraphStyleId { Val = style.StyleName };
        }
    }
}
