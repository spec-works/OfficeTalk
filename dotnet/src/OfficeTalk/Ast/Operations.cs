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

/// <summary>INSERT IMAGE BEFORE/AFTER "source" — insert an image.</summary>
public class InsertImageOperation : Operation
{
    public InsertPosition Position { get; set; }
    public string Source { get; set; } = string.Empty;
    public Dictionary<string, object> Properties { get; set; } = new();
}

/// <summary>INSERT TABLE BEFORE/AFTER rows=N, columns=M — create a new table.</summary>
public class InsertTableOperation : Operation
{
    public InsertPosition Position { get; set; }
    public int Rows { get; set; }
    public int Columns { get; set; }
    public Dictionary<string, object> Properties { get; set; } = new();
}

/// <summary>LINK "url" — create a hyperlink on the addressed element.</summary>
public class LinkOperation : Operation
{
    public string Url { get; set; } = string.Empty;
}

/// <summary>INSERT LIST BEFORE/AFTER [ordered|unordered] — create a list.</summary>
public class InsertListOperation : Operation
{
    public InsertPosition Position { get; set; }
    public ListType ListType { get; set; } = ListType.Unordered;
    public List<ListItem> Items { get; set; } = new();
}

/// <summary>A single item in an INSERT LIST operation.</summary>
public class ListItem
{
    public ContentValue Content { get; set; } = new();
    public bool IsNested { get; set; }
}

/// <summary>SET RUNS — replace element content with individually formatted runs.</summary>
public class SetRunsOperation : Operation
{
    public List<RunDefinition> Runs { get; set; } = new();
}

/// <summary>A single run in a SET RUNS operation.</summary>
public class RunDefinition
{
    public ContentValue Content { get; set; } = new();
    public Dictionary<string, object> Properties { get; set; } = new();
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

/// <summary>Type of list for INSERT LIST operations.</summary>
public enum ListType
{
    Ordered,
    Unordered
}
