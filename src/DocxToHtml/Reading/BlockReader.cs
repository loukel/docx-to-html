using DocumentFormat.OpenXml;
using DocxToHtml.Internal;
using DocxToHtml.Model;

namespace DocxToHtml.Reading;

internal sealed class BlockReader
{
    private readonly ReadingContext context;
    private readonly InlineReader inlineReader;
    private readonly TableReader tableReader;
    private int paragraphCounter;

    public BlockReader(ReadingContext context)
    {
        this.context = context;
        inlineReader = new InlineReader(context, this);
        tableReader = new TableReader(context, this);
    }

    public List<Block> ReadBlocks(OpenXmlElement container)
    {
        var blocks = new List<Block>();
        foreach (var child in MarkupCompatibility.Children(container))
            ReadBlock(child, blocks);

        return blocks;
    }

    private void ReadBlock(OpenXmlElement element, List<Block> blocks)
    {
        switch (element.LocalName)
        {
            case "p":
                blocks.Add(ReadParagraph(element));
                return;
            case "tbl":
                blocks.Add(tableReader.Read(element));
                return;
            case "sdt":
                var content = OpenXmlValues.Child(element, "sdtContent");
                if (content is not null)
                    blocks.AddRange(ReadBlocks(content));
                return;
            case "sectPr":
            case "bookmarkStart":
            case "bookmarkEnd":
            case "proofErr":
            case "permStart":
            case "permEnd":
            case "commentRangeStart":
            case "commentRangeEnd":
                return;
            case "ins":
            case "del":
            case "moveFrom":
            case "moveTo":
                if (RevisionFilter.Include(element.LocalName, context.Options.TrackedChanges))
                    blocks.AddRange(ReadBlocks(element));
                return;
            case "customXml":
            case "smartTag":
                blocks.AddRange(ReadBlocks(element));
                return;
            case "altChunk":
                context.Warnings.Add("UnsupportedContent", "An altChunk part was not imported.", context.Location);
                return;
            case "tblPr":
            case "tblGrid":
            case "trPr":
            case "tcPr":
            case "pPr":
                return;
            default:
                RecoverUnknown(element, blocks);
                return;
        }
    }

    private ParagraphBlock ReadParagraph(OpenXmlElement paragraph)
    {
        paragraphCounter++;
        var previousLocation = context.Location;
        context.Location = previousLocation + "/p[" + paragraphCounter.ToString(CultureInfo.InvariantCulture) + "]";
        try
        {
            var properties = OpenXmlValues.Child(paragraph, "pPr");
            var styleId = OpenXmlValues.Attribute(OpenXmlValues.Child(properties, "pStyle"), "val");
            var definition = context.Styles.Find(styleId);
            var resolved = context.Styles.ResolveParagraph(styleId);
            if (!string.IsNullOrEmpty(styleId) && definition is null)
            {
                context.Warnings.Add("UnknownStyle", $"Paragraph style '{styleId}' was not found.", context.Location);
            }

            var styleName = !string.IsNullOrEmpty(styleId)
                ? definition?.Name ?? styleId
                : resolved.ChainFromSelfToRoot.Count > 0 ? resolved.Name : null;
            var isNormal = string.IsNullOrEmpty(styleId)
                ? resolved.IsNormal || resolved.ChainFromSelfToRoot.Count == 0
                : definition is not null && resolved.IsNormal;

            var decision = HeadingMapper.Decide(
                context.Options,
                styleName,
                isNormal,
                PropertyReader.ReadOutlineLevel(properties),
                resolved.ChainFromSelfToRoot);

            var numbering = PropertyReader.ReadNumbering(properties);
            if (numbering is NumberingSlot.UnsetNumbering)
                numbering = resolved.Numbering;

            var activeNumbering = numbering as NumberingSlot.ActiveNumbering;
            ListInformation? list = null;
            if (activeNumbering is not null && decision.Level is null)
            {
                list = context.Numbering.Describe(
                    activeNumbering.NumberingId,
                    activeNumbering.Level,
                    context.Warnings,
                    context.Location);
            }

            var block = new ParagraphBlock
            {
                StyleId = definition?.StyleId ?? (string.IsNullOrEmpty(styleId) ? resolved.StyleId : styleId),
                StyleName = styleName,
                IsNormalStyle = isNormal,
                Kind = decision.Kind,
                HeadingLevel = decision.Level,
                MappedTag = decision.MappedTag,
                MappedClass = decision.MappedClass,
                List = list,
                Baseline = resolved.BaselineRun,
                DirectFormatting = PropertyReader.ReadParagraph(properties, context.ThemeColors, context.Warnings, context.Location)
            };
            inlineReader.ReadParagraph(paragraph, block);
            if (activeNumbering is not null
                && decision.Level is not null
                && block.Inlines.Any(ParagraphBlock.IsVisibleInline))
            {
                var headingNumber = context.Numbering.TakeLabel(activeNumbering.NumberingId, activeNumbering.Level);
                if (!string.IsNullOrEmpty(headingNumber))
                    block.Inlines.Insert(0, new TextRunInline(headingNumber, block.Baseline));
            }

            return block;
        }
        finally
        {
            context.Location = previousLocation;
        }
    }

    private void RecoverUnknown(OpenXmlElement element, List<Block> blocks)
    {
        var recovered = new List<Block>();
        var foundKnownChild = false;
        foreach (var child in MarkupCompatibility.Children(element))
        {
            if (child.LocalName is "p" or "tbl" or "sdt" or "customXml")
            {
                ReadBlock(child, recovered);
                foundKnownChild = true;
            }
        }

        if (!foundKnownChild)
        {
            var text = VisibleText.From(element, context.Options.TrackedChanges);
            if (!string.IsNullOrWhiteSpace(text))
            {
                recovered.Add(new ParagraphBlock
                {
                    Kind = ParagraphKind.Normal,
                    IsNormalStyle = true,
                    Baseline = context.Styles.DocumentDefaultRun,
                    Inlines = new List<Inline> { new TextRunInline(text, context.Styles.DocumentDefaultRun) }
                });
            }
        }

        if (recovered.Count > 0)
        {
            context.Warnings.Add(
                "UnknownElement",
                $"Unsupported element '{element.LocalName}' was converted and its text was kept.",
                context.Location);
            blocks.AddRange(recovered);
        }
    }
}
