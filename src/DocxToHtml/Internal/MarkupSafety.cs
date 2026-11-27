using System.Net;

namespace DocxToHtml.Internal;

internal static class MarkupSafety
{
    public static string Encode(string? value) => WebUtility.HtmlEncode(value ?? "");

    public static string ClassToken(string token)
    {
        var builder = new StringBuilder(token.Length);
        foreach (var character in token)
        {
            if (char.IsAsciiLetterOrDigit(character) || character is '-' or '_')
                builder.Append(character);
        }

        if (builder.Length == 0)
            return "c";

        if (char.IsDigit(builder[0]))
            return "c-" + builder;

        return builder.ToString();
    }

    public static string Identifier(string value)
    {
        var builder = new StringBuilder(value.Length);
        foreach (var character in value)
        {
            if (char.IsAsciiLetterOrDigit(character) || character is '-' or '_' or ':')
                builder.Append(character);
            else
                builder.Append('-');
        }

        if (builder.Length == 0)
            return "bookmark";

        if (char.IsDigit(builder[0]))
            builder.Insert(0, "id-");

        return builder.ToString();
    }

    public static string LanguageTag(string? language)
    {
        if (string.IsNullOrWhiteSpace(language))
            return "en";

        foreach (var character in language)
        {
            if (!char.IsAsciiLetterOrDigit(character) && character != '-')
                return "en";
        }

        return language;
    }
}
