namespace OfficeTalk.Ast;

/// <summary>
/// Base class for all OfficeTalk operations.
/// </summary>
public abstract class Operation
{
    /// <summary>Source line number.</summary>
    public int Line { get; set; }
}

/// <summary>SET content — replace element content with new text or content block.</summary>
public class SetOperation : Operation
{
    public ContentValue Content { get; set; } = new();
}

/// <summary>REPLACE [ALL] "search" WITH "replacement"</summary>
public class ReplaceOperation : Operation
{
    public string Search { get; set; } = string.Empty;
    public string Replacement { get; set; } = string.Empty;
    public bool IsAll { get; set; }
}

/// <summary>INSERT BEFORE content</summary>
public class InsertBeforeOperation : Operation
{
    public ContentValue Content { get; set; } = new();
}

/// <summary>INSERT AFTER content</summary>
public class InsertAfterOperation : Operation
{
    public ContentValue Content { get; set; } = new();
}

/// <summary>DELETE [ROW|COLUMN|SLIDE|SHEET]</summary>
public class DeleteOperation : Operation
{
    public DeleteTarget Target { get; set; } = DeleteTarget.Element;
}

/// <summary>APPEND content to end of element.</summary>
public class AppendOperation : Operation
{
    public ContentValue Content { get; set; } = new();
}

/// <summary>PREPEND content to beginning of element.</summary>
public class PrependOperation : Operation
{
    public ContentValue Content { get; set; } = new();
}

/// <summary>FORMAT key=value, key=value — apply formatting properties.</summary>
public class FormatOperation : Operation
{
    public Dictionary<string, object> Properties { get; set; } = new();
}

/// <summary>STYLE "styleName" — apply a named style.</summary>
public class StyleOperation : Operation
{
    public string StyleName { get; set; } = string.Empty;
}

/// <summary>INSERT ROW BEFORE/AFTER</summary>
public class InsertRowOperation : Operation
{
    public InsertPosition Position { get; set; }
}

/// <summary>INSERT COLUMN BEFORE/AFTER</summary>
public class InsertColumnOperation : Operation
{
    public InsertPosition Position { get; set; }
}

/// <summary>MERGE CELLS TO address</summary>
public class MergeCellsOperation : Operation
{
    public Address TargetAddress { get; set; } = new();
}

/// <summary>SET CELLS "val1", "val2", ...</summary>
public class SetCellsOperation : Operation
{
    public List<string> Values { get; set; } = new();
}

/// <summary>INSERT SLIDE BEFORE/AFTER</summary>
public class InsertSlideOperation : Operation
{
    public InsertPosition Position { get; set; }
}

/// <summary>DUPLICATE SLIDE</summary>
public class DuplicateSlideOperation : Operation
{
}

/// <summary>RENAME SHEET "newName"</summary>
public class RenameSheetOperation : Operation
{
    public string NewName { get; set; } = string.Empty;
}

/// <summary>ADD SHEET "name"</summary>
public class AddSheetOperation : Operation
{
    public string Name { get; set; } = string.Empty;
}

/// <summary>COMMENT "text" — add a comment anchored to the addressed element.</summary>
public class CommentOperation : Operation
{
    public ContentValue Content { get; set; } = new();
}

/// <summary>
/// INSERT TEXTBOX left=… top=… width=… height=… text="…" [align=…]
/// Creates a floating text box at the specified position and size.
/// </summary>
public class InsertTextboxOperation : Operation
{
    /// <summary>Required text content of the text box.</summary>
    public string Text { get; set; } = string.Empty;

    /// <summary>Horizontal position (e.g. "100pt"). Null if not specified.</summary>
    public string? Left { get; set; }

    /// <summary>Vertical position (e.g. "200pt"). Null if not specified.</summary>
    public string? Top { get; set; }

    /// <summary>Required width of the text box (e.g. "300pt"). Validated as non-null by SyntacticValidator.</summary>
    public string? Width { get; set; }

    /// <summary>Required height of the text box (e.g. "50pt"). Validated as non-null by SyntacticValidator.</summary>
    public string? Height { get; set; }

    /// <summary>Text alignment inside the box: "left", "center", or "right". Null if not specified.</summary>
    public string? Align { get; set; }
}

/// <summary>
/// INSERT SHAPE shapeType left=… top=… width=… height=…
/// Creates a basic drawing shape (e.g. rectangle, oval) at the specified position and size.
/// </summary>
public class InsertShapeOperation : Operation
{
    /// <summary>Shape type identifier (e.g. "rectangle", "oval"). Empty string if not specified.</summary>
    public string ShapeType { get; set; } = string.Empty;

    /// <summary>Horizontal position (e.g. "100pt"). Null if not specified.</summary>
    public string? Left { get; set; }

    /// <summary>Vertical position (e.g. "200pt"). Null if not specified.</summary>
    public string? Top { get; set; }

    /// <summary>Required width of the shape (e.g. "300pt"). Validated as non-null by SyntacticValidator.</summary>
    public string? Width { get; set; }

    /// <summary>Required height of the shape (e.g. "50pt"). Validated as non-null by SyntacticValidator.</summary>
    public string? Height { get; set; }
}

/// <summary>
/// Represents content that can be either an inline string or a content block.
/// </summary>
public class ContentValue
{
    public string Text { get; set; } = string.Empty;
    public bool IsContentBlock { get; set; }

    public ContentValue() { }

    public ContentValue(string text, bool isContentBlock = false)
    {
        Text = text;
        IsContentBlock = isContentBlock;
    }
}

/// <summary>The target of a DELETE operation.</summary>
public enum DeleteTarget
{
    Element,
    Row,
    Column,
    Slide,
    Sheet
}

/// <summary>Position for INSERT operations.</summary>
public enum InsertPosition
{
    Before,
    After
}
