using DocxToHtml.Internal;

namespace DocxToHtml.Reading;

internal static class SymbolMapper
{
    private static readonly Dictionary<(string Font, int Code), string> known = new()
    {
        [("symbol", 0xB7)] = "·",
        [("symbol", 0xF0B7)] = "•",
        [("symbol", 0xA7)] = "§",
        [("wingdings", 0x6C)] = "●",
        [("wingdings", 0xA7)] = "■",
        [("wingdings", 0xD8)] = "➢",
        [("wingdings", 0xF06C)] = "●",
        [("wingdings", 0xF0A7)] = "■",
        [("wingdings", 0xF0D8)] = "➢"
    };

    public static string? Map(string? font, string? characterHex, WarningSink warnings, string? location)
    {
        if (string.IsNullOrEmpty(characterHex)
            || !int.TryParse(characterHex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var codePoint))
        {
            warnings.Add("UnsupportedSymbol", "A symbol character could not be read.", location);
            return null;
        }

        var fontKey = (font ?? "").Trim().ToLowerInvariant();
        if (known.TryGetValue((fontKey, codePoint), out var mapped))
            return mapped;

        if (codePoint is >= 0x20 and <= 0xD7FF or >= 0xE000 and <= 0xFFFD)
        {
            if (codePoint is >= 0xE000 and <= 0xF8FF or >= 0xF000 and <= 0xF0FF)
            {
                warnings.Add(
                    "UnsupportedSymbol",
                    $"No Unicode mapping for symbol font '{font}' character {characterHex}.",
                    location);
                return null;
            }

            return char.ConvertFromUtf32(codePoint);
        }

        warnings.Add(
            "UnsupportedSymbol",
            $"No Unicode mapping for symbol font '{font}' character {characterHex}.",
            location);
        return null;
    }
}
