using System.Net;

namespace DocxToHtml.Internal;

internal static class HtmlText
{
    public static string EncodePreservingSpaces(string? text)
    {
        if (string.IsNullOrEmpty(text))
            return "";

        var encoded = WebUtility.HtmlEncode(text);
        var builder = new StringBuilder(encoded.Length);
        for (var index = 0; index < encoded.Length; index++)
        {
            var character = encoded[index];
            if (character != ' ')
            {
                builder.Append(character);
                continue;
            }

            var adjacentSpace = index == 0
                || index == encoded.Length - 1
                || encoded[index - 1] == ' '
                || encoded[index + 1] == ' ';
            builder.Append(adjacentSpace ? "&#160;" : " ");
        }

        return builder.ToString();
    }

    public static string CollapseWhitespace(string? text)
    {
        if (string.IsNullOrEmpty(text))
            return "";

        var builder = new StringBuilder(text.Length);
        var pendingSpace = false;
        foreach (var character in text)
        {
            if (char.IsWhiteSpace(character) || character == '\u00A0')
            {
                pendingSpace = builder.Length > 0;
                continue;
            }

            if (pendingSpace)
            {
                builder.Append(' ');
                pendingSpace = false;
            }

            builder.Append(character);
        }

        return builder.ToString();
    }
}
