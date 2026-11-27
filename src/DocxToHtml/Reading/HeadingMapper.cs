using DocxToHtml.Model;

namespace DocxToHtml.Reading;

internal readonly record struct HeadingDecision(
    int? Level,
    ParagraphKind Kind,
    string? MappedTag,
    string? MappedClass);

internal static class HeadingMapper
{
    public static HeadingDecision Decide(
        DocxToHtmlOptions options,
        string? styleName,
        bool isNormalStyle,
        int? paragraphOutlineLevel,
        IReadOnlyList<StyleDefinition> chainFromSelfToRoot)
    {
        if (!string.IsNullOrWhiteSpace(styleName)
            && options.StyleMap.TryGetValue(styleName, out var mapping))
        {
            return new HeadingDecision(
                LevelFromTag(mapping.Tag),
                ParagraphKind.Mapped,
                string.IsNullOrWhiteSpace(mapping.Tag) ? "p" : mapping.Tag.Trim().ToLowerInvariant(),
                mapping.CssClass);
        }

        if (!options.UseSemanticHeadings)
            return Plain(isNormalStyle);

        if (TryParseHeadingName(styleName, out var namedLevel))
            return Heading(namedLevel);

        if (IsTitle(styleName))
            return new HeadingDecision(1, ParagraphKind.Title, null, null);

        if (IsSubtitle(styleName))
            return new HeadingDecision(null, ParagraphKind.Subtitle, null, null);

        if (paragraphOutlineLevel is >= 0 and <= 8)
            return Heading(paragraphOutlineLevel.Value + 1);

        foreach (var style in chainFromSelfToRoot)
        {
            if (style.OutlineLevel is >= 0 and <= 8)
                return Heading(style.OutlineLevel.Value + 1);
        }

        for (var index = 1; index < chainFromSelfToRoot.Count; index++)
        {
            var ancestor = chainFromSelfToRoot[index];
            if (TryParseHeadingName(ancestor.Name, out var ancestorLevel))
                return Heading(ancestorLevel);

            if (IsTitle(ancestor.Name))
                return Heading(1);
        }

        return Plain(isNormalStyle);
    }

    public static bool TryParseHeadingName(string? styleName, out int level)
    {
        level = 0;
        if (string.IsNullOrWhiteSpace(styleName))
            return false;

        var trimmed = styleName.Trim();
        if (trimmed.Length is < 8 or > 10)
            return false;

        if (!trimmed.StartsWith("heading", StringComparison.OrdinalIgnoreCase))
            return false;

        var remainder = trimmed[7..].Trim();
        if (remainder.Length != 1 || remainder[0] is < '1' or > '9')
            return false;

        level = remainder[0] - '0';
        return true;
    }

    public static int? LevelFromTag(string? tag)
    {
        if (string.IsNullOrWhiteSpace(tag))
            return null;

        var trimmed = tag.Trim().ToLowerInvariant();
        if (trimmed.Length < 2 || trimmed[0] != 'h')
            return null;

        if (!int.TryParse(trimmed[1..], NumberStyles.Integer, CultureInfo.InvariantCulture, out var level))
            return null;

        return level >= 1 ? level : null;
    }

    private static bool IsTitle(string? styleName) =>
        string.Equals(styleName?.Trim(), "title", StringComparison.OrdinalIgnoreCase);

    private static bool IsSubtitle(string? styleName) =>
        string.Equals(styleName?.Trim(), "subtitle", StringComparison.OrdinalIgnoreCase);

    private static HeadingDecision Heading(int level) =>
        new(level, ParagraphKind.Styled, null, null);

    private static HeadingDecision Plain(bool isNormalStyle) =>
        new(null, isNormalStyle ? ParagraphKind.Normal : ParagraphKind.Styled, null, null);
}
