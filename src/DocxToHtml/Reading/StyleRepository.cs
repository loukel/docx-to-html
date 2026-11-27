using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocxToHtml.Internal;
using DocxToHtml.Model;

namespace DocxToHtml.Reading;

internal sealed class StyleDefinition
{
    public required string StyleId { get; init; }
    public required string Name { get; init; }
    public string? BasedOnStyleId { get; init; }
    public string Type { get; init; } = "paragraph";
    public bool IsDefault { get; init; }
    public int? OutlineLevel { get; init; }
    public NumberingSlot Numbering { get; init; } = NumberingSlot.Unset;
    public ParagraphFormattingPatch Paragraph { get; init; } = ParagraphFormattingPatch.Empty;
    public bool HasTableBorders { get; init; }
}

internal sealed class ResolvedParagraphStyle
{
    public required string StyleId { get; init; }
    public required string Name { get; init; }
    public bool IsNormal { get; init; }
    public ParagraphFormattingPatch ExplicitFormatting { get; init; } = ParagraphFormattingPatch.Empty;
    public RunFormatting BaselineRun { get; init; } = RunFormatting.Empty;
    public NumberingSlot Numbering { get; init; } = NumberingSlot.Unset;
    public IReadOnlyList<StyleDefinition> ChainFromSelfToRoot { get; init; } = Array.Empty<StyleDefinition>();
}

internal sealed class StyleRepository
{
    private readonly Dictionary<string, StyleDefinition> stylesById = new(StringComparer.Ordinal);
    private readonly RunFormatting documentDefaultRun;
    private readonly ParagraphFormattingPatch documentDefaultParagraph;
    private readonly string? defaultParagraphStyleId;

    private StyleRepository(
        RunFormatting documentDefaultRun,
        ParagraphFormattingPatch documentDefaultParagraph,
        string? defaultParagraphStyleId)
    {
        this.documentDefaultRun = documentDefaultRun;
        this.documentDefaultParagraph = documentDefaultParagraph;
        this.defaultParagraphStyleId = defaultParagraphStyleId;
    }

    public RunFormatting DocumentDefaultRun => documentDefaultRun;

    public ParagraphFormattingPatch DocumentDefaultParagraph => documentDefaultParagraph;

    public string? DefaultParagraphStyleId => defaultParagraphStyleId;

    public StyleCatalog Catalog { get; } = new();

    public static StyleRepository Load(MainDocumentPart mainDocumentPart, ThemeColorMap themeColors, WarningSink warnings)
    {
        var stylesRoot = mainDocumentPart.StyleDefinitionsPart?.Styles;
        var defaults = ReadDefaults(stylesRoot, themeColors, warnings);
        var repository = new StyleRepository(defaults.Run, defaults.Paragraph, defaultParagraphStyleId: null);

        if (stylesRoot is not null)
        {
            foreach (var style in stylesRoot.Elements<DocumentFormat.OpenXml.Wordprocessing.Style>())
            {
                var definition = ReadDefinition(style, themeColors, warnings);
                if (definition is null)
                    continue;

                repository.stylesById[definition.StyleId] = definition;
            }
        }

        string? defaultId = repository.stylesById.Values
            .FirstOrDefault(style => style.IsDefault && style.Type.Equals("paragraph", StringComparison.OrdinalIgnoreCase))
            ?.StyleId;

        defaultId ??= repository.stylesById.Values
            .FirstOrDefault(style => style.StyleId.Equals("Normal", StringComparison.OrdinalIgnoreCase))
            ?.StyleId;

        defaultId ??= repository.stylesById.Values
            .FirstOrDefault(style => style.Name.Equals("Normal", StringComparison.OrdinalIgnoreCase))
            ?.StyleId;

        var loaded = new StyleRepository(defaults.Run, defaults.Paragraph, defaultId);
        foreach (var pair in repository.stylesById)
            loaded.stylesById[pair.Key] = pair.Value;

        foreach (var definition in loaded.stylesById.Values)
        {
            if (definition.Type.Equals("paragraph", StringComparison.OrdinalIgnoreCase))
            {
                var resolved = loaded.ResolveParagraph(definition.StyleId);
                loaded.Catalog.AddParagraph(definition.StyleId, resolved.ExplicitFormatting);
            }
            else if (definition.Type.Equals("table", StringComparison.OrdinalIgnoreCase))
            {
                loaded.Catalog.AddTable(
                    definition.StyleId,
                    new TableStyleInfo(definition.Paragraph, definition.HasTableBorders));
            }
        }

        return loaded;
    }

    public ResolvedParagraphStyle ResolveParagraph(string? styleId)
    {
        var effectiveStyleId = string.IsNullOrEmpty(styleId) ? defaultParagraphStyleId : styleId;
        var chain = BuildChain(effectiveStyleId);
        var explicitFormatting = ParagraphFormattingPatch.Empty;
        for (var index = chain.Count - 1; index >= 0; index--)
            explicitFormatting = explicitFormatting.Overlay(chain[index].Paragraph);

        var baseline = explicitFormatting.Run.Apply(documentDefaultRun);
        var numbering = NumberingSlot.Unset;
        foreach (var style in chain)
        {
            if (style.Numbering is not NumberingSlot.UnsetNumbering)
            {
                numbering = style.Numbering;
                break;
            }
        }

        var name = chain.Count > 0 ? chain[0].Name : "Normal";
        var isNormal = chain.Count == 0
            || chain[0].IsDefault
            || name.Equals("Normal", StringComparison.OrdinalIgnoreCase)
            || (effectiveStyleId is not null && effectiveStyleId == defaultParagraphStyleId && name.Equals("Normal", StringComparison.OrdinalIgnoreCase));

        if (chain.Count > 0 && (chain[0].IsDefault || chain[0].Name.Equals("Normal", StringComparison.OrdinalIgnoreCase)))
            isNormal = true;

        return new ResolvedParagraphStyle
        {
            StyleId = effectiveStyleId ?? "",
            Name = name,
            IsNormal = isNormal,
            ExplicitFormatting = explicitFormatting,
            BaselineRun = baseline,
            Numbering = numbering,
            ChainFromSelfToRoot = chain
        };
    }

    public RunFormattingPatch ResolveCharacterPatch(string? characterStyleId)
    {
        var chain = BuildChain(characterStyleId);
        var patch = RunFormattingPatch.Empty;
        for (var index = chain.Count - 1; index >= 0; index--)
            patch = patch.Overlay(chain[index].Paragraph.Run);

        return patch;
    }

    public StyleDefinition? Find(string? styleId)
    {
        if (string.IsNullOrEmpty(styleId))
            return null;

        return stylesById.TryGetValue(styleId, out var definition) ? definition : null;
    }

    private List<StyleDefinition> BuildChain(string? styleId)
    {
        var chain = new List<StyleDefinition>();
        var visited = new HashSet<string>(StringComparer.Ordinal);
        var current = styleId;
        while (!string.IsNullOrEmpty(current) && visited.Add(current) && stylesById.TryGetValue(current, out var definition))
        {
            chain.Add(definition);
            current = definition.BasedOnStyleId;
        }

        return chain;
    }

    private static (RunFormatting Run, ParagraphFormattingPatch Paragraph) ReadDefaults(
        DocumentFormat.OpenXml.Wordprocessing.Styles? stylesRoot,
        ThemeColorMap themeColors,
        WarningSink warnings)
    {
        var defaults = stylesRoot is null ? null : OpenXmlValues.Child(stylesRoot, "docDefaults");
        var runProperties = OpenXmlValues.Child(OpenXmlValues.Child(defaults, "rPrDefault"), "rPr");
        var paragraphProperties = OpenXmlValues.Child(OpenXmlValues.Child(defaults, "pPrDefault"), "pPr");
        var paragraph = PropertyReader.ReadParagraph(paragraphProperties, themeColors, warnings, "styles/docDefaults");
        var runPatch = PropertyReader.ReadRun(runProperties, themeColors, warnings, "styles/docDefaults");
        paragraph = paragraph with { Run = paragraph.Run.Overlay(runPatch) };
        return (runPatch.Apply(RunFormatting.Empty), paragraph);
    }

    private static StyleDefinition? ReadDefinition(
        OpenXmlElement style,
        ThemeColorMap themeColors,
        WarningSink warnings)
    {
        var styleId = OpenXmlValues.Attribute(style, "styleId");
        if (string.IsNullOrEmpty(styleId))
            return null;

        var name = OpenXmlValues.Attribute(OpenXmlValues.Child(style, "name"), "val") ?? styleId;
        var paragraphProperties = OpenXmlValues.Child(style, "pPr");
        var hasBorders = style.Descendants().Any(element => element.LocalName == "tblBorders");
        var paragraphFormatting = PropertyReader.ReadParagraph(paragraphProperties, themeColors, warnings, "style:" + styleId);
        var styleRun = PropertyReader.ReadRun(OpenXmlValues.Child(style, "rPr"), themeColors, warnings, "style:" + styleId);

        return new StyleDefinition
        {
            StyleId = styleId,
            Name = name,
            BasedOnStyleId = OpenXmlValues.Attribute(OpenXmlValues.Child(style, "basedOn"), "val"),
            Type = OpenXmlValues.Attribute(style, "type") ?? "paragraph",
            IsDefault = OpenXmlValues.OnOffAttribute(style, "default") == true,
            OutlineLevel = PropertyReader.ReadOutlineLevel(paragraphProperties),
            Numbering = PropertyReader.ReadNumbering(paragraphProperties),
            Paragraph = paragraphFormatting with { Run = paragraphFormatting.Run.Overlay(styleRun) },
            HasTableBorders = hasBorders
        };
    }
}
