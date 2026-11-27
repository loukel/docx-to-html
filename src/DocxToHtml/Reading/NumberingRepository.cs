using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxToHtml.Internal;
using DocxToHtml.Model;

namespace DocxToHtml.Reading;

internal sealed class NumberingRepository
{
    private readonly Dictionary<int, AbstractNumberingDefinition> definitions = new();

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
        return new LevelDefinition
        {
            Level = level.LevelIndex?.Value ?? 0,
            Start = level.StartNumberingValue?.Val?.Value ?? 1,
            Format = format
        };
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

        public bool TryGetLevel(int level, out LevelDefinition definition)
        {
            if (levels.TryGetValue(level, out definition!))
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
                    Format = pair.Value.Format
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
