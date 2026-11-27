namespace DocxToHtml.Internal;

internal static class StyleSlug
{
    public static string Create(string? name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return "style";

        var builder = new StringBuilder(name.Length);
        var pendingSeparator = false;
        foreach (var character in name.Trim().ToLowerInvariant())
        {
            if (char.IsAsciiLetterOrDigit(character))
            {
                if (pendingSeparator && builder.Length > 0)
                    builder.Append('-');

                pendingSeparator = false;
                builder.Append(character);
            }
            else
            {
                pendingSeparator = true;
            }
        }

        return builder.Length == 0 ? "style" : builder.ToString();
    }
}
