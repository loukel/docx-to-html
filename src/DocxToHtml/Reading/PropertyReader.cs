using DocumentFormat.OpenXml;
using DocxToHtml.Internal;
using DocxToHtml.Model;

namespace DocxToHtml.Reading;

internal static class PropertyReader
{
    public static RunFormattingPatch ReadRun(OpenXmlElement? properties, ThemeColorMap themeColors, WarningSink warnings, string? location)
    {
        if (properties is null)
            return RunFormattingPatch.Empty;

        var strike = OpenXmlValues.ToggleChild(properties, "strike");
        var doubleStrike = OpenXmlValues.ToggleChild(properties, "dstrike");
        bool? strikeValue = strike is null && doubleStrike is null
            ? null
            : (strike ?? false) || (doubleStrike ?? false);

        return new RunFormattingPatch
        {
            Bold = OpenXmlValues.ToggleChild(properties, "b"),
            Italic = OpenXmlValues.ToggleChild(properties, "i"),
            Underline = ReadUnderline(properties),
            Strike = strikeValue,
            VerticalPosition = ReadVerticalPosition(properties),
            ColorHex = ReadColor(properties, themeColors, warnings, location),
            HighlightHex = ReadHighlight(properties, warnings, location),
            FontSizeHalfPoints = ReadHalfPoints(properties),
            FontFamily = ReadFontFamily(properties)
        };
    }

    public static ParagraphFormattingPatch ReadParagraph(OpenXmlElement? properties, ThemeColorMap themeColors, WarningSink warnings, string? location)
    {
        if (properties is null)
            return ParagraphFormattingPatch.Empty;

        var spacing = OpenXmlValues.Child(properties, "spacing");
        var indentation = OpenXmlValues.Child(properties, "ind");
        var lineRule = OpenXmlValues.Attribute(spacing, "lineRule");

        return new ParagraphFormattingPatch
        {
            Alignment = ReadAlignment(properties),
            SpacingBeforeTwips = ReadTwips(spacing, "before"),
            SpacingAfterTwips = ReadTwips(spacing, "after"),
            IndentLeftTwips = ReadTwips(indentation, "left") ?? ReadTwips(indentation, "start"),
            IndentRightTwips = ReadTwips(indentation, "right") ?? ReadTwips(indentation, "end"),
            FirstLineTwips = ReadTwips(indentation, "firstLine"),
            HangingTwips = ReadTwips(indentation, "hanging"),
            LineValue = ReadTwips(spacing, "line"),
            LineRule = string.IsNullOrEmpty(lineRule) ? null : lineRule.ToLowerInvariant(),
            Run = ReadRun(OpenXmlValues.Child(properties, "rPr"), themeColors, warnings, location)
        };
    }

    public static int? ReadOutlineLevel(OpenXmlElement? paragraphProperties)
    {
        return OpenXmlValues.IntegerAttribute(OpenXmlValues.Child(paragraphProperties, "outlineLvl"), "val");
    }

    public static NumberingSlot ReadNumbering(OpenXmlElement? paragraphProperties)
    {
        var numberingProperties = OpenXmlValues.Child(paragraphProperties, "numPr");
        if (numberingProperties is null)
            return NumberingSlot.Unset;

        var numberingId = OpenXmlValues.IntegerAttribute(OpenXmlValues.Child(numberingProperties, "numId"), "val");
        if (numberingId is null or 0)
            return new NumberingSlot.ClearedNumbering();

        var level = OpenXmlValues.IntegerAttribute(OpenXmlValues.Child(numberingProperties, "ilvl"), "val") ?? 0;
        return new NumberingSlot.ActiveNumbering(numberingId.Value, level);
    }

    private static bool? ReadUnderline(OpenXmlElement properties)
    {
        var underline = OpenXmlValues.Child(properties, "u");
        if (underline is null)
            return null;

        var value = OpenXmlValues.Attribute(underline, "val");
        if (string.IsNullOrEmpty(value))
            return true;

        return !value.Equals("none", StringComparison.OrdinalIgnoreCase)
            && !value.Equals("0", StringComparison.OrdinalIgnoreCase)
            && !value.Equals("false", StringComparison.OrdinalIgnoreCase);
    }

    private static VerticalTextPosition? ReadVerticalPosition(OpenXmlElement properties)
    {
        var verticalAlignment = OpenXmlValues.Child(properties, "vertAlign");
        if (verticalAlignment is null)
            return null;

        var value = OpenXmlValues.Attribute(verticalAlignment, "val");
        return value?.ToLowerInvariant() switch
        {
            "superscript" => VerticalTextPosition.Superscript,
            "subscript" => VerticalTextPosition.Subscript,
            _ => VerticalTextPosition.Baseline
        };
    }

    private static string? ReadColor(OpenXmlElement properties, ThemeColorMap themeColors, WarningSink warnings, string? location)
    {
        var color = OpenXmlValues.Child(properties, "color");
        if (color is null)
            return null;

        var literal = OpenXmlValues.Attribute(color, "val");
        if (IsHexColor(literal))
            return literal!.ToUpperInvariant();

        var themeColor = OpenXmlValues.Attribute(color, "themeColor");
        if (string.IsNullOrEmpty(themeColor))
            return null;

        var resolved = themeColors.Resolve(themeColor);
        if (resolved is null)
        {
            warnings.Add("ThemeColor", $"Theme colour '{themeColor}' could not be resolved.", location);
            return null;
        }

        if (!string.IsNullOrEmpty(OpenXmlValues.Attribute(color, "themeTint"))
            || !string.IsNullOrEmpty(OpenXmlValues.Attribute(color, "themeShade")))
        {
            warnings.Add(
                "ThemeColor",
                $"Theme tint or shade on '{themeColor}' was not applied; the base theme colour was used.",
                location);
        }

        return resolved;
    }

    private static string? ReadHighlight(OpenXmlElement properties, WarningSink warnings, string? location)
    {
        var highlight = OpenXmlValues.Child(properties, "highlight");
        var name = OpenXmlValues.Attribute(highlight, "val");
        if (string.IsNullOrEmpty(name) || name.Equals("none", StringComparison.OrdinalIgnoreCase))
            return null;

        var hex = HighlightColors.ToHex(name);
        if (hex is not null)
            return hex;

        warnings.Add("Highlight", $"Highlight colour '{name}' is not supported.", location);
        return null;
    }

    private static int? ReadHalfPoints(OpenXmlElement properties)
    {
        var fontSize = OpenXmlValues.Child(properties, "sz");
        var text = OpenXmlValues.Attribute(fontSize, "val");
        if (text is null)
            return null;

        return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var halfPoints)
            ? halfPoints
            : null;
    }

    private static string? ReadFontFamily(OpenXmlElement properties)
    {
        var fonts = OpenXmlValues.Child(properties, "rFonts");
        var name = OpenXmlValues.Attribute(fonts, "ascii") ?? OpenXmlValues.Attribute(fonts, "hAnsi");
        return string.IsNullOrWhiteSpace(name) ? null : name;
    }

    private static string? ReadAlignment(OpenXmlElement properties)
    {
        var justification = OpenXmlValues.Attribute(OpenXmlValues.Child(properties, "jc"), "val");
        return justification?.ToLowerInvariant() switch
        {
            null or "" => null,
            "center" => "center",
            "right" or "end" => "right",
            "both" or "distribute" or "justify" => "justify",
            "left" or "start" => "left",
            _ => null
        };
    }

    private static int? ReadTwips(OpenXmlElement? element, string localName)
    {
        var text = OpenXmlValues.Attribute(element, localName);
        if (text is null)
            return null;

        if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var twips))
            return twips;

        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
            return (int)Math.Round(number);

        return null;
    }

    private static bool IsHexColor(string? value)
    {
        if (value is null || value.Length != 6)
            return false;

        if (value.Equals("auto", StringComparison.OrdinalIgnoreCase))
            return false;

        foreach (var character in value)
        {
            var isHex = character is >= '0' and <= '9'
                or >= 'a' and <= 'f'
                or >= 'A' and <= 'F';
            if (!isHex)
                return false;
        }

        return true;
    }
}

internal static class HighlightColors
{
    public static string? ToHex(string name) => name.ToLowerInvariant() switch
    {
        "yellow" => "FFFF00",
        "green" => "00FF00",
        "cyan" => "00FFFF",
        "magenta" => "FF00FF",
        "blue" => "0000FF",
        "red" => "FF0000",
        "darkblue" => "00008B",
        "darkcyan" => "008B8B",
        "darkgreen" => "006400",
        "darkmagenta" => "8B008B",
        "darkred" => "8B0000",
        "darkyellow" => "B8860B",
        "darkgray" => "A9A9A9",
        "lightgray" => "D3D3D3",
        "black" => "000000",
        "white" => "FFFFFF",
        _ => null
    };
}
