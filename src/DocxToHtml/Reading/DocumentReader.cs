using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocxToHtml.Internal;
using DocxToHtml.Model;

namespace DocxToHtml.Reading;

internal sealed class DocumentReader
{
    public DocumentModel Read(WordprocessingDocument wordDocument, DocxToHtmlOptions options, string? sourceName)
    {
        var main = wordDocument.MainDocumentPart
            ?? throw new InvalidDataException("The document does not contain a main document part.");
        var body = main.Document?.Body
            ?? throw new InvalidDataException("The document does not contain a Word body.");

        var warnings = new WarningSink(options.OnWarning);
        var themeColors = ThemeColorMap.Load(main);
        var styles = StyleRepository.Load(main, themeColors, warnings);
        var numbering = NumberingRepository.Load(main);
        var context = new ReadingContext
        {
            Options = options,
            Warnings = warnings,
            Styles = styles,
            Numbering = numbering,
            ThemeColors = themeColors,
            Images = new ImageCollector(options, warnings),
            Comments = CommentLookup.Load(wordDocument),
            CurrentPart = main
        };
        var blockReader = new BlockReader(context);
        var bodyBlocks = blockReader.ReadBlocks(body);
        if (context.Fields.HasOpenField)
        {
            warnings.Add("UnclosedField", "A field was still open at the end of the body.", "body");
            context.Fields.Clear();
        }

        var headerBlocks = new List<Block>();
        var footerBlocks = new List<Block>();
        if (options.IncludeHeadersAndFooters)
            ReadHeaderAndFooter(main, body, context, blockReader, headerBlocks, footerBlocks);

        var notes = ReadNotes(main, context, blockReader);
        bodyBlocks = RunMerger.MergeBlocks(bodyBlocks);
        headerBlocks = RunMerger.MergeBlocks(headerBlocks);
        footerBlocks = RunMerger.MergeBlocks(footerBlocks);
        foreach (var note in notes)
            note.Blocks = RunMerger.MergeBlocks(note.Blocks);

        bodyBlocks = ListGrouper.Group(bodyBlocks, numbering, warnings);
        headerBlocks = ListGrouper.Group(headerBlocks, numbering, warnings);
        footerBlocks = ListGrouper.Group(footerBlocks, numbering, warnings);
        foreach (var note in notes)
            note.Blocks = ListGrouper.Group(note.Blocks, numbering, warnings);

        var normal = styles.ResolveParagraph(styles.DefaultParagraphStyleId);
        var bodyParagraph = styles.DocumentDefaultParagraph.Overlay(normal.ExplicitFormatting);
        var language = ReadLanguage(main) ?? "en";

        return new DocumentModel
        {
            Title = ResolveTitle(wordDocument, options, sourceName),
            Language = MarkupSafety.LanguageTag(language),
            BodyBlocks = bodyBlocks,
            HeaderBlocks = headerBlocks,
            FooterBlocks = footerBlocks,
            Notes = notes,
            Images = context.Images.Images,
            Warnings = warnings.Warnings.ToList(),
            Styles = styles.Catalog,
            BodyFontFamily = normal.BaselineRun.FontFamily,
            BodyFontSizeHalfPoints = normal.BaselineRun.FontSizeHalfPoints,
            ParagraphSpacingBeforeTwips = bodyParagraph.SpacingBeforeTwips,
            ParagraphSpacingAfterTwips = bodyParagraph.SpacingAfterTwips,
            DefaultLineValue = bodyParagraph.LineValue,
            DefaultLineRule = bodyParagraph.LineRule
        };
    }

    private static void ReadHeaderAndFooter(
        MainDocumentPart main,
        OpenXmlElement body,
        ReadingContext context,
        BlockReader blockReader,
        List<Block> headerBlocks,
        List<Block> footerBlocks)
    {
        var section = FindFirstSection(body);
        if (section is null)
            return;

        var headerReference = SelectReference(section, "headerReference");
        var footerReference = SelectReference(section, "footerReference");
        if (headerReference is not null)
            ReadReferencedPart(main, headerReference, context, blockReader, headerBlocks, isHeader: true);

        if (footerReference is not null)
            ReadReferencedPart(main, footerReference, context, blockReader, footerBlocks, isHeader: false);
    }

    private static void ReadReferencedPart(
        MainDocumentPart main,
        OpenXmlElement reference,
        ReadingContext context,
        BlockReader blockReader,
        List<Block> destination,
        bool isHeader)
    {
        var relationshipId = OpenXmlValues.Attribute(reference, "id");
        if (string.IsNullOrEmpty(relationshipId))
            return;

        OpenXmlPart? part = null;
        OpenXmlElement? root = null;
        if (isHeader)
        {
            foreach (var headerPart in main.HeaderParts)
            {
                if (main.GetIdOfPart(headerPart) == relationshipId)
                {
                    part = headerPart;
                    root = headerPart.Header;
                    break;
                }
            }
        }
        else
        {
            foreach (var footerPart in main.FooterParts)
            {
                if (main.GetIdOfPart(footerPart) == relationshipId)
                {
                    part = footerPart;
                    root = footerPart.Footer;
                    break;
                }
            }
        }

        if (part is null || root is null)
        {
            context.Warnings.Add(
                "MissingHeaderFooter",
                "A header or footer relationship could not be read.",
                isHeader ? "header" : "footer");
            return;
        }

        var previousPart = context.CurrentPart;
        var previousLocation = context.Location;
        context.CurrentPart = part;
        context.Location = isHeader ? "header" : "footer";
        context.Fields.Clear();
        destination.AddRange(blockReader.ReadBlocks(root));
        context.Fields.Clear();
        context.CurrentPart = previousPart;
        context.Location = previousLocation;
    }

    private static OpenXmlElement? SelectReference(OpenXmlElement section, string localName)
    {
        OpenXmlElement? first = null;
        foreach (var child in MarkupCompatibility.Children(section))
        {
            if (child.LocalName != localName)
                continue;

            first ??= child;
            var type = OpenXmlValues.Attribute(child, "type");
            if (type is null || type.Equals("default", StringComparison.OrdinalIgnoreCase))
                return child;
        }

        return first;
    }

    private static OpenXmlElement? FindFirstSection(OpenXmlElement body)
    {
        foreach (var element in body.Descendants())
        {
            if (element.LocalName == "sectPr")
                return element;
        }

        return null;
    }

    private static List<NoteModel> ReadNotes(MainDocumentPart main, ReadingContext context, BlockReader blockReader)
    {
        var notes = new List<NoteModel>();
        if (!context.Options.IncludeFootnotes)
            return notes;

        AppendNotes(main.FootnotesPart, main.FootnotesPart?.Footnotes, isEndnote: false);
        AppendNotes(main.EndnotesPart, main.EndnotesPart?.Endnotes, isEndnote: true);
        return notes;

        void AppendNotes(OpenXmlPart? part, OpenXmlElement? root, bool isEndnote)
        {
            if (part is null || root is null)
                return;

            var previousPart = context.CurrentPart;
            var previousLocation = context.Location;
            context.CurrentPart = part;
            context.Location = isEndnote ? "endnotes" : "footnotes";
            context.Fields.Clear();
            foreach (var request in context.Notes.Where(note => note.IsEndnote == isEndnote))
            {
                var note = FindNote(root, request.Identifier);
                if (note is null || IsSeparator(note))
                {
                    context.Warnings.Add(
                        "MissingNote",
                        $"Note {request.Identifier} was referenced but is missing or is a separator.",
                        context.Location);
                    continue;
                }

                notes.Add(new NoteModel
                {
                    DisplayNumber = request.DisplayNumber,
                    IsEndnote = isEndnote,
                    Blocks = blockReader.ReadBlocks(note)
                });
            }

            context.Fields.Clear();
            context.CurrentPart = previousPart;
            context.Location = previousLocation;
        }
    }

    private static OpenXmlElement? FindNote(OpenXmlElement root, int identifier)
    {
        foreach (var child in root.ChildElements)
        {
            if (child.LocalName is not ("footnote" or "endnote"))
                continue;

            if (OpenXmlValues.IntegerAttribute(child, "id") == identifier)
                return child;
        }

        return null;
    }

    private static bool IsSeparator(OpenXmlElement note)
    {
        var identifier = OpenXmlValues.IntegerAttribute(note, "id") ?? 0;
        if (identifier <= 0)
            return true;

        var type = OpenXmlValues.Attribute(note, "type");
        return type is not null && type.ToLowerInvariant() is "separator" or "continuationseparator" or "continuationnotice";
    }

    private static string? ReadLanguage(MainDocumentPart main)
    {
        var defaults = OpenXmlValues.Child(main.StyleDefinitionsPart?.Styles, "docDefaults");
        var runProperties = OpenXmlValues.Child(OpenXmlValues.Child(defaults, "rPrDefault"), "rPr");
        return OpenXmlValues.Attribute(OpenXmlValues.Child(runProperties, "lang"), "val");
    }

    private static string ResolveTitle(WordprocessingDocument wordDocument, DocxToHtmlOptions options, string? sourceName)
    {
        if (!string.IsNullOrWhiteSpace(options.DocumentTitle))
            return options.DocumentTitle.Trim();

        var packageTitle = wordDocument.PackageProperties.Title;
        if (!string.IsNullOrWhiteSpace(packageTitle))
            return packageTitle.Trim();

        if (!string.IsNullOrWhiteSpace(sourceName))
            return sourceName.Trim();

        return "Document";
    }
}
