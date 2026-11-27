using DocumentFormat.OpenXml;
using DocxToHtml.Internal;
using DocxToHtml.Model;

namespace DocxToHtml.Reading;

internal sealed class InlineReader
{
    private static readonly HashSet<string> silentElements = new(StringComparer.Ordinal)
    {
        "proofErr", "bookmarkEnd", "commentRangeStart", "commentRangeEnd",
        "permStart", "permEnd", "moveFromRangeStart", "moveFromRangeEnd",
        "moveToRangeStart", "moveToRangeEnd", "customXmlInsRangeStart", "customXmlInsRangeEnd",
        "customXmlDelRangeStart", "customXmlDelRangeEnd", "lastRenderedPageBreak",
        "pPr", "rPr", "sectPr", "fldChar"
    };

    private readonly ReadingContext context;
    private readonly ImageReader imageReader;

    public InlineReader(ReadingContext context, BlockReader blockReader)
    {
        this.context = context;
        imageReader = new ImageReader(context, blockReader);
    }

    public void ReadParagraph(OpenXmlElement paragraph, ParagraphBlock block)
    {
        var cursor = new InlineCursor { Paragraph = block };
        ReadChildren(paragraph, block.Inlines, block.Baseline, cursor);
    }

    private void ReadChildren(OpenXmlElement parent, List<Inline> destination, RunFormatting formatting, InlineCursor cursor)
    {
        foreach (var child in MarkupCompatibility.Children(parent))
            ReadChild(child, destination, formatting, cursor);
    }

    private void ReadChild(OpenXmlElement child, List<Inline> destination, RunFormatting formatting, InlineCursor cursor)
    {
        switch (child.LocalName)
        {
            case "r":
                ReadRun(child, destination, formatting, cursor);
                return;
            case "hyperlink":
                ReadHyperlink(child, destination, formatting, cursor);
                return;
            case "fldSimple":
                ReadSimpleField(child, destination, formatting, cursor);
                return;
            case "sdt":
                var content = OpenXmlValues.Child(child, "sdtContent") ?? child;
                ReadChildren(content, destination, formatting, cursor);
                return;
            case "smartTag":
            case "customXml":
                ReadChildren(child, destination, formatting, cursor);
                return;
            case "ins":
            case "del":
            case "moveFrom":
            case "moveTo":
                if (RevisionFilter.Include(child.LocalName, context.Options.TrackedChanges))
                    ReadChildren(child, destination, formatting, cursor);
                return;
            case "bookmarkStart":
                ReadBookmark(child, destination, cursor);
                return;
            case "drawing":
            case "pict":
            case "object":
                var before = destination.Count;
                imageReader.ReadContainer(child, destination);
                NoteContent(destination, before, cursor);
                return;
            default:
                if (!silentElements.Contains(child.LocalName) && child.ChildElements.Count > 0)
                {
                    var text = VisibleText.From(child, context.Options.TrackedChanges);
                    if (!string.IsNullOrEmpty(text))
                    {
                        context.Warnings.Add(
                            "UnknownElement",
                            $"Unsupported element '{child.LocalName}' was kept as text.",
                            context.Location);
                        Emit(new TextRunInline(text, formatting), destination, cursor);
                    }
                }
                return;
        }
    }

    private void ReadRun(OpenXmlElement run, List<Inline> destination, RunFormatting paragraphBaseline, InlineCursor cursor)
    {
        var formatting = ResolveFormatting(OpenXmlValues.Child(run, "rPr"), paragraphBaseline);
        foreach (var child in MarkupCompatibility.Children(run))
        {
            switch (child.LocalName)
            {
                case "t":
                    Emit(new TextRunInline(child.InnerText, formatting), destination, cursor);
                    break;
                case "delText":
                    if (context.Options.TrackedChanges == TrackedChangesMode.Reject)
                        Emit(new TextRunInline(child.InnerText, formatting), destination, cursor);
                    break;
                case "br":
                    var breakType = OpenXmlValues.Attribute(child, "type");
                    Emit(
                        breakType is not null && (breakType.Equals("page", StringComparison.OrdinalIgnoreCase)
                            || breakType.Equals("column", StringComparison.OrdinalIgnoreCase))
                            ? new PageBreakInline()
                            : new LineBreakInline(),
                        destination,
                        cursor);
                    break;
                case "cr":
                    Emit(new LineBreakInline(), destination, cursor);
                    break;
                case "tab":
                case "ptab":
                    Emit(new TabInline(), destination, cursor);
                    break;
                case "noBreakHyphen":
                    Emit(new NonBreakingHyphenInline(), destination, cursor);
                    break;
                case "softHyphen":
                    Emit(new SoftHyphenInline(), destination, cursor);
                    break;
                case "sym":
                    var mapped = SymbolMapper.Map(
                        OpenXmlValues.Attribute(child, "font"),
                        OpenXmlValues.Attribute(child, "char"),
                        context.Warnings,
                        context.Location);
                    if (!string.IsNullOrEmpty(mapped))
                        Emit(new SymbolInline(mapped), destination, cursor);
                    break;
                case "footnoteReference":
                    EmitNote(child, isEndnote: false, destination, cursor);
                    break;
                case "endnoteReference":
                    EmitNote(child, isEndnote: true, destination, cursor);
                    break;
                case "footnoteRef":
                case "endnoteRef":
                    break;
                case "fldChar":
                    ReadFieldChar(child);
                    break;
                case "instrText":
                case "delInstrText":
                    context.Fields.AppendInstruction(child.InnerText);
                    break;
                case "drawing":
                case "pict":
                case "object":
                    var before = destination.Count;
                    imageReader.ReadContainer(child, destination);
                    NoteContent(destination, before, cursor);
                    break;
                case "commentReference":
                    ReadComment(child, destination, cursor);
                    break;
                case "lastRenderedPageBreak":
                case "rPr":
                case "proofErr":
                    break;
                default:
                    if (child.LocalName is "ins" or "del" or "moveFrom" or "moveTo")
                    {
                        if (RevisionFilter.Include(child.LocalName, context.Options.TrackedChanges))
                            ReadChildren(child, destination, formatting, cursor);
                        break;
                    }

                    var text = VisibleText.From(child, context.Options.TrackedChanges);
                    if (!string.IsNullOrEmpty(text) && !silentElements.Contains(child.LocalName))
                    {
                        context.Warnings.Add(
                            "UnknownElement",
                            $"Unsupported run content '{child.LocalName}' was kept as text.",
                            context.Location);
                        Emit(new TextRunInline(text, formatting), destination, cursor);
                    }
                    break;
            }
        }
    }

    private void ReadHyperlink(OpenXmlElement hyperlink, List<Inline> destination, RunFormatting formatting, InlineCursor cursor)
    {
        var children = new List<Inline>();
        var nested = new InlineCursor();
        ReadChildren(hyperlink, children, formatting, nested);
        var href = ResolveHyperlink(hyperlink);
        if (href is null)
        {
            context.Warnings.Add("UnresolvedHyperlink", "A hyperlink could not be resolved and its text was kept.", context.Location);
            foreach (var child in children)
                Emit(child, destination, cursor);
            return;
        }

        Emit(new HyperlinkInline { Href = href, Children = children }, destination, cursor);
        if (!string.IsNullOrEmpty(nested.ElementId) && cursor.Paragraph is not null && cursor.AtStart && cursor.Paragraph.ElementId is null)
            cursor.Paragraph.ElementId = nested.ElementId;
    }

    private void ReadSimpleField(OpenXmlElement field, List<Inline> destination, RunFormatting formatting, InlineCursor cursor)
    {
        var instruction = OpenXmlValues.Attribute(field, "instr") ?? "";
        string? href = null;
        if (HyperlinkInstruction.TryParse(instruction, out var parsed))
            href = parsed;

        context.Fields.PushResult(href);
        try
        {
            ReadChildren(field, destination, formatting, cursor);
        }
        finally
        {
            context.Fields.Pop();
        }
    }

    private void ReadFieldChar(OpenXmlElement fieldChar)
    {
        var type = OpenXmlValues.Attribute(fieldChar, "fldCharType");
        if (type is null)
            return;

        if (type.Equals("begin", StringComparison.OrdinalIgnoreCase))
            context.Fields.Begin();
        else if (type.Equals("separate", StringComparison.OrdinalIgnoreCase))
            context.Fields.Separate();
        else if (type.Equals("end", StringComparison.OrdinalIgnoreCase))
            context.Fields.End();
    }

    private void ReadBookmark(OpenXmlElement bookmark, List<Inline> destination, InlineCursor cursor)
    {
        var name = OpenXmlValues.Attribute(bookmark, "name");
        if (string.IsNullOrWhiteSpace(name) || name.StartsWith('_'))
            return;

        var identifier = MarkupSafety.Identifier(name);
        if (cursor.AtStart && cursor.Paragraph is not null && cursor.Paragraph.ElementId is null)
        {
            cursor.Paragraph.ElementId = identifier;
            return;
        }

        destination.Add(new BookmarkInline(identifier));
    }

    private void ReadComment(OpenXmlElement commentReference, List<Inline> destination, InlineCursor cursor)
    {
        if (!context.Options.IncludeComments)
            return;

        var identifier = OpenXmlValues.Attribute(commentReference, "id");
        if (string.IsNullOrEmpty(identifier))
            return;

        var text = context.Comments.Find(identifier);
        if (text is null)
        {
            context.Warnings.Add("MissingComment", $"Comment {identifier} was not found.", context.Location);
            return;
        }

        if (text.Length > 0)
            Emit(new CommentInline(text), destination, cursor);
    }

    private void EmitNote(OpenXmlElement reference, bool isEndnote, List<Inline> destination, InlineCursor cursor)
    {
        if (!context.Options.IncludeFootnotes)
            return;

        var identifier = OpenXmlValues.IntegerAttribute(reference, "id");
        if (identifier is null)
            return;

        var displayNumber = context.AssignNoteNumber(identifier.Value, isEndnote);
        Emit(new NoteReferenceInline(displayNumber, isEndnote), destination, cursor);
    }

    private void Emit(Inline inline, List<Inline> destination, InlineCursor cursor)
    {
        if (context.Fields.ShouldSkipText)
            return;

        var href = context.Fields.CurrentHyperlinkHref;
        if (!string.IsNullOrEmpty(href) && inline is not BookmarkInline and not HyperlinkInline)
        {
            destination.Add(new HyperlinkInline
            {
                Href = href,
                Children = new List<Inline> { inline }
            });
        }
        else
        {
            destination.Add(inline);
        }

        if (ParagraphBlock.IsVisibleInline(inline))
            cursor.AtStart = false;
    }

    private static void NoteContent(List<Inline> destination, int before, InlineCursor cursor)
    {
        for (var index = before; index < destination.Count; index++)
        {
            if (ParagraphBlock.IsVisibleInline(destination[index]))
            {
                cursor.AtStart = false;
                return;
            }
        }
    }

    private RunFormatting ResolveFormatting(OpenXmlElement? runProperties, RunFormatting baseline)
    {
        var characterStyleId = OpenXmlValues.Attribute(OpenXmlValues.Child(runProperties, "rStyle"), "val");
        if (context.Options.DirectFormatting == DirectFormattingMode.None)
        {
            if (string.IsNullOrEmpty(characterStyleId))
                return baseline;

            return context.Styles.ResolveCharacterPatch(characterStyleId).Apply(baseline);
        }

        var patch = string.IsNullOrEmpty(characterStyleId)
            ? RunFormattingPatch.Empty
            : context.Styles.ResolveCharacterPatch(characterStyleId);
        if (runProperties is not null)
            patch = patch.Overlay(PropertyReader.ReadRun(runProperties, context.ThemeColors, context.Warnings, context.Location));

        return patch.Apply(baseline);
    }

    private string? ResolveHyperlink(OpenXmlElement hyperlink)
    {
        string? href = null;
        var relationshipId = OpenXmlValues.Attribute(hyperlink, "id");
        if (!string.IsNullOrEmpty(relationshipId))
        {
            var relationship = context.CurrentPart.HyperlinkRelationships
                .FirstOrDefault(candidate => candidate.Id == relationshipId);
            if (relationship is null)
                context.Warnings.Add("UnresolvedHyperlink", $"Hyperlink relationship '{relationshipId}' was not found.", context.Location);
            else
                href = relationship.Uri.IsAbsoluteUri ? relationship.Uri.AbsoluteUri : relationship.Uri.OriginalString;
        }

        var anchor = OpenXmlValues.Attribute(hyperlink, "anchor");
        if (!string.IsNullOrEmpty(anchor))
        {
            var fragment = "#" + anchor;
            href = string.IsNullOrEmpty(href) ? fragment : href + fragment;
        }

        return string.IsNullOrEmpty(href) ? null : href;
    }

    private sealed class InlineCursor
    {
        public ParagraphBlock? Paragraph { get; init; }
        public bool AtStart { get; set; } = true;
        public string? ElementId => Paragraph?.ElementId;
    }
}
