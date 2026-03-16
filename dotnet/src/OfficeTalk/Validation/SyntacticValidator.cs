using OfficeTalk.Ast;

namespace OfficeTalk.Validation;

/// <summary>
/// Validates the syntactic structure of a parsed OfficeTalk AST.
/// Checks for required fields, valid operation combinations, and structural correctness.
/// </summary>
public class SyntacticValidator
{
    /// <summary>
    /// Validate the AST structure.
    /// </summary>
    public ValidationResult Validate(OfficeTalkDocument document)
    {
        var result = new ValidationResult();

        // Check version
        if (string.IsNullOrWhiteSpace(document.Version))
        {
            result.Errors.Add(new ValidationDiagnostic(
                ValidationCategory.Syntax,
                "Missing or empty OFFICETALK version header."));
        }

        // Check for parse errors
        foreach (var error in document.Errors)
        {
            result.Errors.Add(new ValidationDiagnostic(
                ValidationCategory.Syntax,
                error.Message,
                error.Line,
                error.Column));
        }

        // Check for mixed read/write operations
        bool hasWriteOps = document.OperationBlocks.Count > 0 || document.PropertySettings.Count > 0;
        bool hasInspectOps = document.InspectBlocks.Count > 0;

        if (hasWriteOps && hasInspectOps)
        {
            result.Errors.Add(new ValidationDiagnostic(
                ValidationCategory.MixedOperations,
                "Document mixes INSPECT and write operations (AT/PROPERTY). A document must contain either INSPECT blocks or write operations, not both."));
        }

        // Validate operation blocks
        foreach (var block in document.OperationBlocks)
        {
            ValidateBlock(block, document.DocType, result);
        }

        // Validate inspect blocks
        foreach (var block in document.InspectBlocks)
        {
            ValidateInspectBlock(block, result);
        }

        // Validate property settings
        foreach (var prop in document.PropertySettings)
        {
            ValidateProperty(prop, result);
        }

        return result;
    }

    private static void ValidateBlock(OperationBlock block, DocType docType, ValidationResult result)
    {
        // Block must have an address
        if (block.Address.Segments.Count == 0)
        {
            result.Errors.Add(new ValidationDiagnostic(
                ValidationCategory.Syntax,
                "Operation block has an empty address.",
                block.Line));
        }

        // Block must have at least one operation
        if (block.Operations.Count == 0)
        {
            result.Errors.Add(new ValidationDiagnostic(
                ValidationCategory.Syntax,
                "Operation block has no operations.",
                block.Line));
        }

        // Validate individual operations
        foreach (var operation in block.Operations)
        {
            ValidateOperation(operation, docType, result);
        }

        // Check for conflicting operations
        CheckConflicts(block, result);
    }

    private static void ValidateOperation(Operation operation, DocType docType, ValidationResult result)
    {
        switch (operation)
        {
            case SetOperation set:
                if (string.IsNullOrEmpty(set.Content.Text) && !set.Content.IsContentBlock)
                {
                    result.Errors.Add(new ValidationDiagnostic(
                        ValidationCategory.InvalidValue,
                        "SET operation has empty content.",
                        operation.Line));
                }
                break;

            case ReplaceOperation replace:
                if (string.IsNullOrEmpty(replace.Search))
                {
                    result.Errors.Add(new ValidationDiagnostic(
                        ValidationCategory.InvalidValue,
                        "REPLACE operation has empty search string.",
                        operation.Line));
                }
                break;

            case StyleOperation style:
                if (string.IsNullOrWhiteSpace(style.StyleName))
                {
                    result.Errors.Add(new ValidationDiagnostic(
                        ValidationCategory.InvalidValue,
                        "STYLE operation has empty style name.",
                        operation.Line));
                }
                break;

            case InsertSlideOperation when docType != DocType.PowerPoint:
                result.Errors.Add(new ValidationDiagnostic(
                    ValidationCategory.InvalidOperation,
                    "INSERT SLIDE is only valid for PowerPoint documents.",
                    operation.Line));
                break;

            case DuplicateSlideOperation when docType != DocType.PowerPoint:
                result.Errors.Add(new ValidationDiagnostic(
                    ValidationCategory.InvalidOperation,
                    "DUPLICATE SLIDE is only valid for PowerPoint documents.",
                    operation.Line));
                break;

            case RenameSheetOperation when docType != DocType.Excel:
                result.Errors.Add(new ValidationDiagnostic(
                    ValidationCategory.InvalidOperation,
                    "RENAME SHEET is only valid for Excel documents.",
                    operation.Line));
                break;

            case AddSheetOperation when docType != DocType.Excel:
                result.Errors.Add(new ValidationDiagnostic(
                    ValidationCategory.InvalidOperation,
                    "ADD SHEET is only valid for Excel documents.",
                    operation.Line));
                break;

            case CommentOperation comment:
                if (string.IsNullOrEmpty(comment.Content.Text) && !comment.Content.IsContentBlock)
                {
                    result.Errors.Add(new ValidationDiagnostic(
                        ValidationCategory.InvalidValue,
                        "COMMENT operation has empty text.",
                        operation.Line));
                }
                break;
        }
    }

    private static void CheckConflicts(OperationBlock block, ValidationResult result)
    {
        bool hasSet = block.Operations.Any(o => o is SetOperation);
        bool hasDelete = block.Operations.Any(o => o is DeleteOperation);
        bool hasStyle = block.Operations.Any(o => o is StyleOperation);
        bool hasFormat = block.Operations.Any(o => o is FormatOperation);

        if (hasSet && hasDelete)
        {
            result.Errors.Add(new ValidationDiagnostic(
                ValidationCategory.Conflict,
                "Operation block contains both SET and DELETE, which conflict.",
                block.Line));
        }

        if (hasStyle && hasFormat)
        {
            result.Warnings.Add(new ValidationDiagnostic(
                ValidationCategory.StyleOverride,
                "Operation block contains both STYLE and FORMAT. FORMAT properties may override style properties.",
                block.Line));
        }
    }

    private static void ValidateInspectBlock(InspectBlock block, ValidationResult result)
    {
        if (block.Address.Segments.Count == 0)
        {
            result.Errors.Add(new ValidationDiagnostic(
                ValidationCategory.Syntax,
                "INSPECT block has an empty address.",
                block.Line));
        }

        if (block.Depth < 0)
        {
            result.Errors.Add(new ValidationDiagnostic(
                ValidationCategory.InvalidValue,
                $"DEPTH must be a non-negative integer, got {block.Depth}.",
                block.Line));
        }

        if (block.Context < 0)
        {
            result.Errors.Add(new ValidationDiagnostic(
                ValidationCategory.InvalidValue,
                $"CONTEXT must be a non-negative integer, got {block.Context}.",
                block.Line));
        }
    }

    private static void ValidateProperty(PropertySetting prop, ValidationResult result)
    {
        var validNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "title", "author", "subject", "description", "keywords", "category"
        };

        if (!validNames.Contains(prop.Name))
        {
            result.Warnings.Add(new ValidationDiagnostic(
                ValidationCategory.InvalidValue,
                $"Unknown document property: '{prop.Name}'.",
                prop.Line));
        }
    }
}
