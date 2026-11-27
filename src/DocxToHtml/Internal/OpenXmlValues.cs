using DocumentFormat.OpenXml;

namespace DocxToHtml.Internal;

internal static class OpenXmlValues
{
    public static string? Attribute(OpenXmlElement? element, string localName)
    {
        if (element is null)
            return null;

        foreach (var attribute in element.GetAttributes())
        {
            if (string.Equals(attribute.LocalName, localName, StringComparison.Ordinal))
                return attribute.Value;
        }

        return null;
    }

    public static OpenXmlElement? Child(OpenXmlElement? parent, string localName)
    {
        if (parent is null)
            return null;

        foreach (var child in parent.ChildElements)
        {
            if (string.Equals(child.LocalName, localName, StringComparison.Ordinal))
                return child;
        }

        return null;
    }

    public static int? IntegerAttribute(OpenXmlElement? element, string localName)
    {
        var text = Attribute(element, localName);
        if (text is null)
            return null;

        return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var number)
            ? number
            : null;
    }

    public static long? LongAttribute(OpenXmlElement? element, string localName)
    {
        var text = Attribute(element, localName);
        if (text is null)
            return null;

        return long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var number)
            ? number
            : null;
    }

    /// <summary>
    /// Word toggle elements such as <c>w:b</c> are on when the element is present and <c>w:val</c> is omitted.
    /// <c>w:val="0"</c>, <c>false</c>, or <c>off</c> turns the toggle off so a style can be cancelled.
    /// </summary>
    public static bool? ToggleChild(OpenXmlElement? parent, string localName)
    {
        var child = Child(parent, localName);
        if (child is null)
            return null;

        var value = Attribute(child, "val");
        if (string.IsNullOrEmpty(value))
            return true;

        return IsEnabled(value);
    }

    public static bool? OnOffAttribute(OpenXmlElement? element, string localName)
    {
        var value = Attribute(element, localName);
        if (string.IsNullOrEmpty(value))
            return null;

        return IsEnabled(value);
    }

    public static bool IsEnabled(string value)
    {
        return value.ToLowerInvariant() is not ("0" or "false" or "off" or "none");
    }
}
