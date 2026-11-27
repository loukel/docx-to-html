namespace DocxToHtml.Writing;

/// <summary>
/// Indents block elements in generated HTML. Text and inline markup stay on the
/// same line as the element that contains them.
/// </summary>
internal static class HtmlFormatter
{
    private static readonly HashSet<string> AlwaysBreakElements = new(StringComparer.OrdinalIgnoreCase)
    {
        "html", "head", "body", "header", "footer", "section", "nav", "article", "main",
        "div", "ul", "ol", "table", "thead", "tbody", "tfoot", "tr", "colgroup",
        "blockquote", "figure", "figcaption"
    };

    private static readonly HashSet<string> BlockElements = new(StringComparer.OrdinalIgnoreCase)
    {
        "html", "head", "body", "header", "footer", "section", "nav", "article", "main",
        "div", "ul", "ol", "li", "table", "thead", "tbody", "tfoot", "tr", "td", "th",
        "colgroup", "p", "h1", "h2", "h3", "h4", "h5", "h6", "blockquote", "figure", "figcaption",
        "style", "script"
    };

    private static readonly HashSet<string> VoidElements = new(StringComparer.OrdinalIgnoreCase)
    {
        "area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "source", "track", "wbr"
    };

    private static readonly HashSet<string> RawElements = new(StringComparer.OrdinalIgnoreCase)
    {
        "style", "script", "pre", "textarea"
    };

    public static string Format(string html)
    {
        if (string.IsNullOrEmpty(html))
            return html;

        var tokens = Tokenize(html);
        var formatted = new StringBuilder();
        WriteChildren(tokens, 0, tokens.Count, formatted, indentLevel: 0, firstBlockContinuesLine: false);
        return formatted.ToString();
    }

    private static void WriteChildren(
        List<HtmlToken> tokens,
        int start,
        int end,
        StringBuilder formatted,
        int indentLevel,
        bool firstBlockContinuesLine)
    {
        var index = start;
        var continuesLine = firstBlockContinuesLine;
        while (index < end)
        {
            var token = tokens[index];
            if (token.Kind == HtmlTokenKind.Text)
            {
                if (!IsFormattingWhitespace(token.Source))
                {
                    if (!continuesLine)
                        AppendIndent(formatted, indentLevel);
                    formatted.Append(token.Source);
                    if (!continuesLine)
                        formatted.Append('\n');
                    continuesLine = false;
                }

                index++;
                continue;
            }

            if (token.Kind == HtmlTokenKind.End)
            {
                index++;
                continue;
            }

            index = WriteElement(tokens, index, formatted, indentLevel, indentOpening: !continuesLine);
            continuesLine = false;
        }
    }

    private static int WriteElement(List<HtmlToken> tokens, int start, StringBuilder formatted, int indentLevel, bool indentOpening)
    {
        var opening = tokens[start];
        if (opening.Kind == HtmlTokenKind.Doctype)
        {
            formatted.Append(opening.Source).Append('\n');
            return start + 1;
        }

        if (opening.Kind == HtmlTokenKind.Void)
        {
            if (indentOpening)
                AppendIndent(formatted, indentLevel);
            formatted.Append(opening.Source).Append('\n');
            return start + 1;
        }

        var closing = FindClosingToken(tokens, start);
        if (closing < 0)
        {
            if (indentOpening)
                AppendIndent(formatted, indentLevel);
            formatted.Append(opening.Source).Append('\n');
            return start + 1;
        }

        if (RawElements.Contains(opening.Name))
        {
            if (indentOpening)
                AppendIndent(formatted, indentLevel);
            formatted.Append(opening.Source).Append('\n');
            AppendRawContents(tokens, start + 1, closing, formatted, indentLevel + 1);
            AppendIndent(formatted, indentLevel);
            formatted.Append(tokens[closing].Source).Append('\n');
            return closing + 1;
        }

        var breakChildren = AlwaysBreakElements.Contains(opening.Name) || ContainsBlockElement(tokens, start + 1, closing);
        if (!breakChildren)
        {
            if (indentOpening)
                AppendIndent(formatted, indentLevel);
            formatted.Append(opening.Source);
            AppendInlineContents(tokens, start + 1, closing, formatted);
            formatted.Append(tokens[closing].Source).Append('\n');
            return closing + 1;
        }

        if (indentOpening)
            AppendIndent(formatted, indentLevel);
        formatted.Append(opening.Source);
        var childIndex = start + 1;
        var keptLeadingText = false;
        if (!AlwaysBreakElements.Contains(opening.Name))
        {
            while (childIndex < closing && IsPhrasing(tokens, childIndex))
            {
                if (tokens[childIndex].Kind == HtmlTokenKind.Text && IsFormattingWhitespace(tokens[childIndex].Source))
                {
                    childIndex++;
                    continue;
                }

                AppendPhrasing(tokens, ref childIndex, formatted);
                keptLeadingText = true;
            }
        }

        // A newline here would become a space in the text, so "Fruit" must stay glued to a nested list.
        if (!keptLeadingText)
            formatted.Append('\n');
        WriteChildren(tokens, childIndex, closing, formatted, indentLevel + 1, firstBlockContinuesLine: keptLeadingText);
        AppendIndent(formatted, indentLevel);
        formatted.Append(tokens[closing].Source).Append('\n');
        return closing + 1;
    }

    private static void AppendInlineContents(List<HtmlToken> tokens, int start, int end, StringBuilder formatted)
    {
        for (var index = start; index < end; index++)
        {
            if (tokens[index].Kind == HtmlTokenKind.Text && IsFormattingWhitespace(tokens[index].Source))
                continue;

            formatted.Append(tokens[index].Source);
        }
    }

    private static void AppendPhrasing(List<HtmlToken> tokens, ref int index, StringBuilder formatted)
    {
        var token = tokens[index];
        if (token.Kind is HtmlTokenKind.Text or HtmlTokenKind.Void)
        {
            formatted.Append(token.Source);
            index++;
            return;
        }

        var closing = FindClosingToken(tokens, index);
        if (closing < 0)
        {
            formatted.Append(token.Source);
            index++;
            return;
        }

        for (var cursor = index; cursor <= closing; cursor++)
            formatted.Append(tokens[cursor].Source);

        index = closing + 1;
    }

    private static void AppendRawContents(List<HtmlToken> tokens, int start, int end, StringBuilder formatted, int indentLevel)
    {
        var contents = new StringBuilder();
        for (var index = start; index < end; index++)
            contents.Append(tokens[index].Source);

        var text = contents.ToString().Replace("\r\n", "\n", StringComparison.Ordinal).Trim('\n');
        if (text.Length == 0)
            return;

        foreach (var line in text.Split('\n'))
        {
            if (line.Length == 0)
            {
                formatted.Append('\n');
                continue;
            }

            AppendIndent(formatted, indentLevel);
            formatted.Append(line).Append('\n');
        }
    }

    private static bool ContainsBlockElement(List<HtmlToken> tokens, int start, int end)
    {
        for (var index = start; index < end; index++)
        {
            var token = tokens[index];
            if (token.Kind is HtmlTokenKind.Start or HtmlTokenKind.Void && BlockElements.Contains(token.Name))
                return true;
        }

        return false;
    }

    private static bool IsPhrasing(List<HtmlToken> tokens, int index)
    {
        var token = tokens[index];
        return token.Kind switch
        {
            HtmlTokenKind.Text or HtmlTokenKind.Void => true,
            HtmlTokenKind.Start => !BlockElements.Contains(token.Name),
            _ => false
        };
    }

    private static int FindClosingToken(List<HtmlToken> tokens, int start)
    {
        var depth = 0;
        for (var index = start; index < tokens.Count; index++)
        {
            var token = tokens[index];
            if (token.Kind == HtmlTokenKind.Start)
                depth++;
            else if (token.Kind == HtmlTokenKind.End)
            {
                depth--;
                if (depth == 0)
                    return index;
            }
        }

        return -1;
    }

    private static bool IsFormattingWhitespace(string source) =>
        source.Contains('\n') && string.IsNullOrWhiteSpace(source);

    private static void AppendIndent(StringBuilder formatted, int indentLevel) =>
        formatted.Append(' ', indentLevel * 2);

    private static List<HtmlToken> Tokenize(string html)
    {
        var tokens = new List<HtmlToken>();
        var index = 0;
        while (index < html.Length)
        {
            if (html[index] != '<')
            {
                var textStart = index;
                while (index < html.Length && html[index] != '<')
                    index++;

                tokens.Add(new HtmlToken(HtmlTokenKind.Text, "", html[textStart..index]));
                continue;
            }

            var tagStart = index;
            var tagEnd = ReadTagEnd(html, index);
            var source = html[tagStart..(tagEnd + 1)];
            index = tagEnd + 1;

            if (source.StartsWith("<!", StringComparison.Ordinal))
            {
                tokens.Add(new HtmlToken(HtmlTokenKind.Doctype, "", source));
                continue;
            }

            var isClosing = source.StartsWith("</", StringComparison.Ordinal);
            var name = ReadTagName(source, isClosing ? 2 : 1);
            if (isClosing)
                tokens.Add(new HtmlToken(HtmlTokenKind.End, name, source));
            else if (VoidElements.Contains(name) || source.EndsWith("/>", StringComparison.Ordinal))
                tokens.Add(new HtmlToken(HtmlTokenKind.Void, name, source));
            else
                tokens.Add(new HtmlToken(HtmlTokenKind.Start, name, source));
        }

        return tokens;
    }

    private static int ReadTagEnd(string html, int start)
    {
        var quote = '\0';
        for (var index = start; index < html.Length; index++)
        {
            var character = html[index];
            if (quote != '\0')
            {
                if (character == quote)
                    quote = '\0';
                continue;
            }

            if (character is '"' or '\'')
            {
                quote = character;
                continue;
            }

            if (character == '>')
                return index;
        }

        return html.Length - 1;
    }

    private static string ReadTagName(string source, int start)
    {
        var index = start;
        while (index < source.Length && (char.IsLetterOrDigit(source[index]) || source[index] is '-' or ':'))
            index++;

        return source[start..index];
    }

    private enum HtmlTokenKind
    {
        Text,
        Start,
        End,
        Void,
        Doctype
    }

    private readonly record struct HtmlToken(HtmlTokenKind Kind, string Name, string Source);
}
