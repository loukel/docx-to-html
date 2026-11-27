using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxToHtml.Internal;
using DocxToHtml.Model;

namespace DocxToHtml.Reading;

internal sealed class NumberingRepository
{
    private readonly Dictionary<int, AbstractNumberingDefinition> definitions = new();
    private readonly Dictionary<int, int[]> countersByNumberingId = new();

    public static NumberingRepository Load(MainDocumentPart mainDocumentPart)
    {
        var repository = new NumberingRepository();
        var numbering = mainDocumentPart.NumberingDefinitionsPart?.Numbering;
        if (numbering is null)
            return repository;

        var abstracts = new Dictionary<int, AbstractNumberingDefinition>();
        foreach (var abstractNumbering in numbering.Elements<AbstractNum>())
        {
            var abstractId = abstractNumbering.AbstractNumberId?.Value;
            if (abstractId is null)
                continue;

            abstracts[abstractId.Value] = ReadAbstract(abstractNumbering);
        }

        foreach (var instance in numbering.Elements<NumberingInstance>())
        {
            var numberingId = instance.NumberID?.Value;
            var abstractId = instance.AbstractNumId?.Val?.Value;
            if (numberingId is null || abstractId is null || !abstracts.TryGetValue(abstractId.Value, out var definition))
                continue;

            var copy = definition.Clone();
            foreach (var levelOverride in instance.Elements<LevelOverride>())
            {
                var levelIndex = levelOverride.LevelIndex?.Value ?? 0;
                var startOverride = levelOverride.StartOverrideNumberingValue?.Val?.Value;
                if (startOverride is not null)
                    copy.SetStart(levelIndex, startOverride.Value);

                if (levelOverride.Level is not null)
                    copy.SetLevel(ReadLevel(levelOverride.Level));
            }

            repository.definitions[numberingId.Value] = copy;
        }

        return repository;
    }

    public ListInformation? Describe(int numberingId, int level, WarningSink warnings, string? location)
    {
        if (!definitions.TryGetValue(numberingId, out var definition))
        {
            warnings.Add("UnresolvedNumbering", $"Numbering definition {numberingId} was not found. The list was treated as bullets.", location);
            return new ListInformation(numberingId, level, IsBullet: true, FormatClass: null, Start: null);
        }

        if (!definition.TryGetLevel(level, out var levelDefinition))
        {
            warnings.Add("UnresolvedNumbering", $"Numbering level {level} was not found for list {numberingId}.", location);
            return new ListInformation(numberingId, level, IsBullet: true, FormatClass: null, Start: null);
        }

        var format = levelDefinition.Format ?? "decimal";
        var isBullet = format.Equals("bullet", StringComparison.OrdinalIgnoreCase)
            || format.Equals("none", StringComparison.OrdinalIgnoreCase);
        var formatClass = isBullet ? null : NumberFormatClasses.ToClass(format);
        int? start = levelDefinition.Start > 1 ? levelDefinition.Start : null;
        return new ListInformation(numberingId, level, isBullet, formatClass, start);
    }

    /// <summary>
    /// Advances the counters for one numbered heading and returns the label, including the
    /// separator before the heading text. Heading numbers stay in the heading instead of
    /// becoming a nested list.
    /// </summary>
    public string? TakeLabel(int numberingId, int level)
    {
        if (!definitions.TryGetValue(numberingId, out var definition))
            return null;

        if (!definition.TryGetExactLevel(level, out var current))
            return null;

        if (string.Equals(current.Format, "none", StringComparison.OrdinalIgnoreCase))
            return null;

        if (!countersByNumberingId.TryGetValue(numberingId, out var counters))
        {
            counters = new int[9];
            countersByNumberingId[numberingId] = counters;
        }

        var clampedLevel = Math.Clamp(level, 0, counters.Length - 1);
        for (var deeperLevel = clampedLevel + 1; deeperLevel < counters.Length; deeperLevel++)
            counters[deeperLevel] = 0;

        if (counters[clampedLevel] <= 0)
            counters[clampedLevel] = Math.Max(current.Start, 1);
        else
            counters[clampedLevel]++;

        if (string.Equals(current.Format, "bullet", StringComparison.OrdinalIgnoreCase))
        {
            var bullet = string.IsNullOrEmpty(current.LevelText) ? "•" : current.LevelText;
            return bullet + current.Separator;
        }

        var pattern = string.IsNullOrEmpty(current.LevelText)
            ? "%" + (clampedLevel + 1).ToString(CultureInfo.InvariantCulture)
            : current.LevelText;
        var label = pattern;
        for (var placeholder = 9; placeholder >= 1; placeholder--)
        {
            var token = "%" + placeholder.ToString(CultureInfo.InvariantCulture);
            if (!label.Contains(token, StringComparison.Ordinal))
                continue;

            var format = current.LegalNumbering
                ? "decimal"
                : definition.TryGetExactLevel(placeholder - 1, out var referencedLevel)
                    ? referencedLevel.Format ?? "decimal"
                    : "decimal";
            label = label.Replace(token, FormatNumber(counters[placeholder - 1], format), StringComparison.Ordinal);
        }

        if (string.IsNullOrWhiteSpace(label))
            return null;

        return label + current.Separator;
    }

    private static AbstractNumberingDefinition ReadAbstract(AbstractNum abstractNumbering)
    {
        var definition = new AbstractNumberingDefinition();
        foreach (var level in abstractNumbering.Elements<Level>())
            definition.SetLevel(ReadLevel(level));

        return definition;
    }

    private static LevelDefinition ReadLevel(Level level)
    {
        var format = OpenXmlValues.Attribute(level.NumberingFormat, "val") ?? "decimal";
        var suffix = OpenXmlValues.Attribute(OpenXmlValues.Child(level, "suff"), "val");
        return new LevelDefinition
        {
            Level = level.LevelIndex?.Value ?? 0,
            Start = level.StartNumberingValue?.Val?.Value ?? 1,
            Format = format,
            LevelText = OpenXmlValues.Attribute(OpenXmlValues.Child(level, "lvlText"), "val"),
            LegalNumbering = OpenXmlValues.ToggleChild(level, "isLgl") == true,
            Separator = suffix is not null && suffix.Equals("nothing", StringComparison.OrdinalIgnoreCase) ? "" : " "
        };
    }

    private static string FormatNumber(int value, string format)
    {
        if (value < 0)
            value = 0;

        return format.ToLowerInvariant() switch
        {
            "lowerletter" => ToLetters(value, uppercase: false),
            "upperletter" => ToLetters(value, uppercase: true),
            "lowerroman" => ToRoman(value).ToLowerInvariant(),
            "upperroman" => ToRoman(value),
            "decimalzero" => value.ToString("00", CultureInfo.InvariantCulture),
            "bullet" or "none" => "",
            _ => value.ToString(CultureInfo.InvariantCulture)
        };
    }

    private static string ToLetters(int value, bool uppercase)
    {
        if (value <= 0)
            return "0";

        var characters = new List<char>();
        var remaining = value;
        while (remaining > 0)
        {
            remaining--;
            characters.Add((char)((uppercase ? 'A' : 'a') + remaining % 26));
            remaining /= 26;
        }

        characters.Reverse();
        return new string(characters.ToArray());
    }

    private static string ToRoman(int value)
    {
        if (value <= 0)
            return "0";

        var numerals = new (int Amount, string Symbol)[]
        {
            (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
            (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
            (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")
        };
        var roman = new StringBuilder();
        var remaining = value;
        foreach (var (amount, symbol) in numerals)
        {
            while (remaining >= amount)
            {
                roman.Append(symbol);
                remaining -= amount;
            }
        }

        return roman.ToString();
    }

    private sealed class AbstractNumberingDefinition
    {
        private readonly Dictionary<int, LevelDefinition> levels = new();

        public void SetLevel(LevelDefinition level) => levels[level.Level] = level;

        public void SetStart(int level, int start)
        {
            if (!levels.TryGetValue(level, out var existing))
            {
                levels[level] = new LevelDefinition { Level = level, Start = start, Format = "decimal" };
                return;
            }

            existing.Start = start;
        }

        public bool TryGetExactLevel(int level, out LevelDefinition definition)
        {
            if (levels.TryGetValue(level, out definition!))
                return true;

            definition = null!;
            return false;
        }

        public bool TryGetLevel(int level, out LevelDefinition definition)
        {
            if (TryGetExactLevel(level, out definition))
                return true;

            if (levels.TryGetValue(0, out definition!))
                return true;

            definition = null!;
            return false;
        }

        public AbstractNumberingDefinition Clone()
        {
            var copy = new AbstractNumberingDefinition();
            foreach (var pair in levels)
            {
                copy.levels[pair.Key] = new LevelDefinition
                {
                    Level = pair.Value.Level,
                    Start = pair.Value.Start,
                    Format = pair.Value.Format,
                    LevelText = pair.Value.LevelText,
                    LegalNumbering = pair.Value.LegalNumbering,
                    Separator = pair.Value.Separator
                };
            }

            return copy;
        }
    }

    private sealed class LevelDefinition
    {
        public int Level { get; init; }
        public int Start { get; set; }
        public string? Format { get; init; }
        public string? LevelText { get; init; }
        public bool LegalNumbering { get; init; }
        public string Separator { get; init; } = " ";
    }
}

internal static class NumberFormatClasses
{
    public static string? ToClass(string format) => format.ToLowerInvariant() switch
    {
        "decimal" or "bullet" or "none" => null,
        "lowerletter" => "lower-alpha",
        "upperletter" => "upper-alpha",
        "lowerroman" => "lower-roman",
        "upperroman" => "upper-roman",
        "decimalzero" => "decimal-leading-zero",
        "ordinal" => "ordinal",
        "cardinaltext" => "cardinal-text",
        "ordinaltext" => "ordinal-text",
        var other => StyleSlug.Create(other)
    };
}
