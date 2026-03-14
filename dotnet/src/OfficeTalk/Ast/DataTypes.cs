namespace OfficeTalk.Ast;

/// <summary>
/// Represents a color value (hex or named).
/// </summary>
public class ColorValue
{
    public string Raw { get; set; } = string.Empty;
    public byte R { get; set; }
    public byte G { get; set; }
    public byte B { get; set; }
    public byte? A { get; set; }
    public bool IsNamed { get; set; }

    public ColorValue() { }

    public ColorValue(string raw)
    {
        Raw = raw;
        if (raw.StartsWith('#'))
        {
            var hex = raw[1..];
            if (hex.Length == 6)
            {
                R = Convert.ToByte(hex[..2], 16);
                G = Convert.ToByte(hex[2..4], 16);
                B = Convert.ToByte(hex[4..6], 16);
            }
            else if (hex.Length == 8)
            {
                R = Convert.ToByte(hex[..2], 16);
                G = Convert.ToByte(hex[2..4], 16);
                B = Convert.ToByte(hex[4..6], 16);
                A = Convert.ToByte(hex[6..8], 16);
            }
        }
        else
        {
            IsNamed = true;
        }
    }

    public string ToHex() => A.HasValue
        ? $"#{R:X2}{G:X2}{B:X2}{A:X2}"
        : $"#{R:X2}{G:X2}{B:X2}";

    public override string ToString() => IsNamed ? Raw : ToHex();
}

/// <summary>
/// Represents a length value with a unit.
/// </summary>
public class LengthValue
{
    public double Amount { get; set; }
    public LengthUnit Unit { get; set; }
    public string Raw { get; set; } = string.Empty;

    public LengthValue() { }

    public LengthValue(double amount, LengthUnit unit)
    {
        Amount = amount;
        Unit = unit;
        Raw = $"{amount}{UnitSuffix(unit)}";
    }

    public override string ToString() => Raw;

    private static string UnitSuffix(LengthUnit unit) => unit switch
    {
        LengthUnit.Points => "pt",
        LengthUnit.Inches => "in",
        LengthUnit.Centimeters => "cm",
        LengthUnit.Percentage => "%",
        LengthUnit.Emu => "emu",
        _ => ""
    };
}

/// <summary>
/// Length unit types supported by OfficeTalk.
/// </summary>
public enum LengthUnit
{
    Points,
    Inches,
    Centimeters,
    Percentage,
    Emu
}
