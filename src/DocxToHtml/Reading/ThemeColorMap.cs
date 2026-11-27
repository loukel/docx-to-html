using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocxToHtml.Internal;

namespace DocxToHtml.Reading;

internal sealed class ThemeColorMap
{
    private readonly Dictionary<string, string> colors = new(StringComparer.OrdinalIgnoreCase);

    public static ThemeColorMap Load(MainDocumentPart mainDocumentPart)
    {
        var map = new ThemeColorMap();
        var theme = mainDocumentPart.ThemePart?.Theme;
        if (theme is null)
            return map;

        foreach (var element in theme.Descendants())
        {
            if (!IsSchemeColor(element.LocalName))
                continue;

            var hex = ReadSchemeColor(element);
            if (hex is not null)
                map.colors[element.LocalName] = hex;
        }

        return map;
    }

    public string? Resolve(string themeColorName)
    {
        var key = themeColorName.ToLowerInvariant() switch
        {
            "dark1" => "dk1",
            "light1" => "lt1",
            "dark2" => "dk2",
            "light2" => "lt2",
            "hyperlink" => "hlink",
            "followedhyperlink" => "folHlink",
            var other => other
        };

        return colors.TryGetValue(key, out var hex) ? hex : null;
    }

    private static bool IsSchemeColor(string localName) => localName.ToLowerInvariant() is
        "dk1" or "lt1" or "dk2" or "lt2"
        or "accent1" or "accent2" or "accent3" or "accent4" or "accent5" or "accent6"
        or "hlink" or "folhlink";

    private static string? ReadSchemeColor(OpenXmlElement element)
    {
        foreach (var child in element.ChildElements)
        {
            if (child.LocalName.Equals("srgbClr", StringComparison.OrdinalIgnoreCase))
            {
                var value = OpenXmlValues.Attribute(child, "val");
                if (value is { Length: 6 })
                    return value.ToUpperInvariant();
            }

            if (child.LocalName.Equals("sysClr", StringComparison.OrdinalIgnoreCase))
            {
                var lastColor = OpenXmlValues.Attribute(child, "lastClr");
                if (lastColor is { Length: 6 })
                    return lastColor.ToUpperInvariant();
            }
        }

        return null;
    }
}
