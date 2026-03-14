namespace OfficeTalk.Ast;

/// <summary>
/// A predicate used to filter elements within an address segment.
/// </summary>
public abstract class Predicate
{
}

/// <summary>
/// A positional predicate using 1-based indexing (e.g., paragraph[3]).
/// </summary>
public class PositionalPredicate : Predicate
{
    public int Position { get; set; }

    public PositionalPredicate() { }
    public PositionalPredicate(int position) => Position = position;

    public override string ToString() => Position.ToString();
}

/// <summary>
/// A key-value predicate with a comparison operator (e.g., level=2, text~="pattern").
/// </summary>
public class KeyValuePredicate : Predicate
{
    public string Key { get; set; } = string.Empty;
    public PredicateOperator Operator { get; set; }
    public string Value { get; set; } = string.Empty;

    public KeyValuePredicate() { }

    public KeyValuePredicate(string key, PredicateOperator op, string value)
    {
        Key = key;
        Operator = op;
        Value = value;
    }

    public override string ToString()
    {
        var opStr = Operator switch
        {
            PredicateOperator.Equals => "=",
            PredicateOperator.TildeEquals => "~=",
            PredicateOperator.CaretEquals => "^=",
            PredicateOperator.DollarEquals => "$=",
            PredicateOperator.AsteriskEquals => "*=",
            _ => "="
        };
        return $"{Key}{opStr}\"{Value}\"";
    }
}

/// <summary>
/// A bare string predicate, used as a shorthand name (e.g., sheet["Revenue"]).
/// </summary>
public class BareStringPredicate : Predicate
{
    public string Value { get; set; } = string.Empty;

    public BareStringPredicate() { }
    public BareStringPredicate(string value) => Value = value;

    public override string ToString() => $"\"{Value}\"";
}

/// <summary>
/// A cell reference predicate for Excel addresses (e.g., A1, A1:D10).
/// </summary>
public class CellRefPredicate : Predicate
{
    public string CellRef { get; set; } = string.Empty;

    public CellRefPredicate() { }
    public CellRefPredicate(string cellRef) => CellRef = cellRef;

    public override string ToString() => CellRef;
}

/// <summary>
/// Comparison operators used in key-value predicates.
/// </summary>
public enum PredicateOperator
{
    /// <summary>Exact match (=)</summary>
    Equals,
    /// <summary>Regex/I-Regexp match (~=)</summary>
    TildeEquals,
    /// <summary>Starts with (^=)</summary>
    CaretEquals,
    /// <summary>Ends with ($=)</summary>
    DollarEquals,
    /// <summary>Contains (*=)</summary>
    AsteriskEquals
}
