using DocxToHtml.Internal;
using DocxToHtml.Model;
using DocxToHtml.Reading;

namespace DocxToHtml.Writing;

internal sealed class HtmlDocumentWriter
{
    public ConversionResult Write(DocumentModel document, DocxToHtmlOptions options)
    {
        var registry = new ClassRegistry(options);
        var usedStyles = new List<UsedStyle>();
        var headings = new List<CollectedHeading>();
        var htmlBuilder = new StringBuilder();
        var bodyWriter = new BodyWriter(document, options, registry, usedStyles, headings, htmlBuilder);
        bodyWriter.WriteContents();
        var stylesheet = new CssBuilder(document, options, registry, usedStyles).Build();
        var html = options.OutputMode == HtmlOutputMode.BodyFragment
            ? htmlBuilder.ToString()
            : WrapDocument(document, options, htmlBuilder.ToString(), stylesheet);
        html = HtmlFormatter.Format(html);

        return new ConversionResult
        {
            Html = html,
            Css = stylesheet,
            Images = document.Images,
            Warnings = document.Warnings,
            Headings = headings
        };
    }

    private static string WrapDocument(DocumentModel document, DocxToHtmlOptions options, string body, string stylesheet)
    {
        var builder = new StringBuilder();
        builder.Append("<!DOCTYPE html>\n<html lang=\"").Append(MarkupSafety.Encode(document.Language)).Append("\">\n<head>\n");
        builder.Append("<meta charset=\"utf-8\">\n");
        builder.Append("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n");
        builder.Append("<title>").Append(MarkupSafety.Encode(document.Title)).Append("</title>\n");
        if (options.CssMode == CssMode.External)
        {
            builder.Append("<link rel=\"stylesheet\" href=\"")
                .Append(MarkupSafety.Encode(CssFileName(options)))
                .Append("\">\n");
        }
        else if (options.CssMode == CssMode.Embedded)
        {
            builder.Append("<style>\n").Append(stylesheet.Replace("</", "<\\/", StringComparison.OrdinalIgnoreCase)).Append("</style>\n");
        }

        builder.Append("</head>\n<body>\n");
        builder.Append(body);
        if (body.Length > 0 && !body.EndsWith('\n'))
            builder.Append('\n');
        builder.Append("</body>\n</html>\n");
        return builder.ToString();
    }

    internal static string CssFileName(DocxToHtmlOptions options)
    {
        var name = string.IsNullOrWhiteSpace(options.CssFileName) ? "styles.css" : Path.GetFileName(options.CssFileName);
        return string.IsNullOrWhiteSpace(name) ? "styles.css" : name;
    }

    internal static string ImageFolder(DocxToHtmlOptions options)
    {
        if (string.IsNullOrWhiteSpace(options.ImageFolder))
            return "images";

        var parts = options.ImageFolder.Replace('\\', '/').Split('/', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0 || parts.Any(part => part is "." or ".." || part.IndexOfAny([':', '*', '?', '"', '<', '>', '|']) >= 0))
            return "images";

        return string.Join('/', parts);
    }
}

internal sealed class BodyWriter
{
    private readonly DocumentModel document;
    private readonly DocxToHtmlOptions options;
    private readonly ClassRegistry registry;
    private readonly List<UsedStyle> usedStyles;
    private readonly List<CollectedHeading> headings;
    private readonly StringBuilder html;
    private readonly HashSet<string> rememberedStyles = new(StringComparer.Ordinal);

    public BodyWriter(
        DocumentModel document,
        DocxToHtmlOptions options,
        ClassRegistry registry,
        List<UsedStyle> usedStyles,
        List<CollectedHeading> headings,
        StringBuilder html)
    {
        this.document = document;
        this.options = options;
        this.registry = registry;
        this.usedStyles = usedStyles;
        this.headings = headings;
        this.html = html;
    }

    public void WriteContents()
    {
        if (document.HeaderBlocks.Count > 0)
        {
            html.Append("<header class=\"").Append(FeatureClass("doc-header", "color: #444444;", "margin-bottom: 1.5em;")).Append("\">\n");
            WriteBlocks(document.HeaderBlocks);
            html.Append("</header>\n");
        }

        WriteBlocks(document.BodyBlocks);
        WriteNotes(endnotes: false);
        WriteNotes(endnotes: true);

        if (document.FooterBlocks.Count > 0)
        {
            html.Append("<footer class=\"").Append(FeatureClass("doc-footer", "color: #444444;", "margin-top: 1.5em;")).Append("\">\n");
            WriteBlocks(document.FooterBlocks);
            html.Append("</footer>\n");
        }
    }

    private void WriteBlocks(IReadOnlyList<Block> blocks)
    {
        foreach (var block in blocks)
            WriteBlock(block);
    }

    private void WriteBlock(Block block)
    {
        switch (block)
        {
            case ParagraphBlock paragraph:
                WriteParagraph(paragraph);
                break;
            case TableBlock table:
                WriteTable(table);
                break;
            case TextBoxBlock textBox:
                WriteTextBox(textBox.Blocks);
                break;
            case ListBlock list:
                WriteList(list);
                break;
        }
    }

    private void WriteParagraph(ParagraphBlock paragraph)
    {
        if (IsHeadingTag(TagFor(paragraph)) && !HasVisibleInlines(paragraph.Inlines))
        {
            WriteBookmarkAnchors(paragraph);
            return;
        }

        if (!options.IncludeEmptyParagraphs && !paragraph.HasVisibleContent())
            return;

        var segments = SplitTextBoxes(paragraph.Inlines);
        if (segments.Count == 1 && segments[0].TextBox is null)
        {
            WriteParagraphElement(paragraph, segments[0].Inlines);
            return;
        }

        foreach (var segment in segments)
        {
            if (segment.TextBox is not null)
                WriteTextBox(segment.TextBox.Blocks);
            else if (segment.Inlines.Count > 0 || (options.IncludeEmptyParagraphs && paragraph.HasVisibleContent()))
                WriteParagraphElement(paragraph, segment.Inlines);
        }
    }

    private static bool HasVisibleInlines(IReadOnlyList<Inline> inlines) =>
        inlines.Any(ParagraphBlock.IsVisibleInline);

    private static bool IsStructuralListItem(ListItemModel item) =>
        (item.Paragraph is null || !HasVisibleInlines(item.Paragraph.Inlines))
        && item.NestedLists.Count > 0;

    private void WriteBookmarkAnchors(ParagraphBlock paragraph)
    {
        if (!string.IsNullOrEmpty(paragraph.ElementId))
            html.Append("<a id=\"").Append(MarkupSafety.Encode(paragraph.ElementId)).Append("\"></a>");

        foreach (var inline in paragraph.Inlines)
        {
            if (inline is BookmarkInline bookmark)
                html.Append("<a id=\"").Append(MarkupSafety.Encode(bookmark.Identifier)).Append("\"></a>");
        }
    }

    private void WriteParagraphElement(ParagraphBlock paragraph, IReadOnlyList<Inline> inlines)
    {
        var tag = TagFor(paragraph);
        var classes = ClassesFor(paragraph);
        html.Append('<').Append(tag);
        AppendClass(classes);
        if (!string.IsNullOrEmpty(paragraph.ElementId))
            html.Append(" id=\"").Append(MarkupSafety.Encode(paragraph.ElementId)).Append('"');
        html.Append('>');
        WriteInlines(inlines, paragraph.Baseline, allowTextBox: false);
        html.Append("</").Append(tag).Append(">\n");
        CollectHeading(paragraph, tag, inlines);
    }

    private void WriteList(ListBlock list)
    {
        var tag = list.IsBullet ? "ul" : "ol";
        var classes = new List<string>();
        if (!string.IsNullOrEmpty(list.FormatClass))
            classes.Add(FeatureClass(list.FormatClass, "list-style-type: " + list.FormatClass + ";"));
        // An empty parent item is only there so a deeper level has somewhere to sit.
        // Without this, that item paints a bullet and then the real sub-list paints more.
        if (list.Items.Count > 0 && list.Items.All(IsStructuralListItem))
        {
            classes.Add(FeatureClass(
                "list-plain",
                "list-style: none;",
                "margin: 0;",
                "padding: 0;"));
        }

        html.Append('<').Append(tag);
        AppendClass(classes);
        if (!list.IsBullet && list.Start is int start && start != 1)
            html.Append(" start=\"").Append(start.ToString(CultureInfo.InvariantCulture)).Append('"');
        html.Append(">\n");

        foreach (var item in list.Items)
        {
            var paragraph = item.Paragraph;
            var heading = paragraph is not null && IsHeadingTag(TagFor(paragraph));
            var itemClasses = new List<string>();
            if (IsStructuralListItem(item))
            {
                itemClasses.Add(FeatureClass(
                    "marker-none",
                    "list-style: none;"));
            }
            else if (paragraph is not null && !heading)
            {
                itemClasses.AddRange(ClassesFor(paragraph));
            }

            html.Append("<li");
            AppendClass(itemClasses);
            if (paragraph is not null && !heading && !string.IsNullOrEmpty(paragraph.ElementId))
                html.Append(" id=\"").Append(MarkupSafety.Encode(paragraph.ElementId)).Append('"');
            html.Append('>');
            if (paragraph is not null)
            {
                if (heading && HasVisibleInlines(paragraph.Inlines))
                    WriteParagraphElement(paragraph, paragraph.Inlines);
                else if (heading)
                    WriteBookmarkAnchors(paragraph);
                else
                    WriteInlines(paragraph.Inlines, paragraph.Baseline, allowTextBox: true);
            }

            foreach (var nested in item.NestedLists)
                WriteList(nested);

            html.Append("</li>\n");
        }

        html.Append("</").Append(tag).Append(">\n");
    }

    private void WriteTable(TableBlock table)
    {
        var classes = new List<string>();
        if (!string.IsNullOrWhiteSpace(table.StyleName))
        {
            var className = Prefix("t-" + StyleSlug.Create(table.StyleName));
            classes.Add(className);
            Remember(table.StyleId, className, isTable: true);
        }

        if (table.FirstRowLook)
            classes.Add(Prefix("t-first-row"));
        if (table.BandedRows)
            classes.Add(Prefix("t-banded"));

        html.Append("<table");
        AppendClass(classes);
        html.Append(">\n");
        if (table.ColumnWidthPercentages.Count > 0)
        {
            html.Append("<colgroup>\n");
            foreach (var percent in table.ColumnWidthPercentages)
            {
                var className = registry.Register("w-" + percent.ToString(CultureInfo.InvariantCulture), "width: " + percent.ToString(CultureInfo.InvariantCulture) + "%;");
                html.Append("<col class=\"").Append(className).Append("\">\n");
            }
            html.Append("</colgroup>\n");
        }

        var leadingHeaders = 0;
        while (leadingHeaders < table.Rows.Count && table.Rows[leadingHeaders].IsHeader)
            leadingHeaders++;

        var rowsRemainingInSpan = new int[table.ColumnWidthPercentages.Count];
        if (leadingHeaders > 0)
        {
            html.Append("<thead>\n");
            for (var index = 0; index < leadingHeaders; index++)
                WriteRow(table.Rows[index], table.ColumnWidthPercentages, rowsRemainingInSpan);
            html.Append("</thead>\n");
        }

        if (leadingHeaders < table.Rows.Count)
        {
            html.Append("<tbody>\n");
            for (var index = leadingHeaders; index < table.Rows.Count; index++)
                WriteRow(table.Rows[index], table.ColumnWidthPercentages, rowsRemainingInSpan);
            html.Append("</tbody>\n");
        }

        html.Append("</table>\n");
    }

    private void WriteRow(TableRowModel row, IReadOnlyList<int> columnWidths, int[] rowsRemainingInSpan)
    {
        html.Append("<tr>\n");
        var columnIndex = 0;
        foreach (var cell in row.Cells)
        {
            while (columnIndex < rowsRemainingInSpan.Length && rowsRemainingInSpan[columnIndex] > 0)
                columnIndex++;

            var tag = row.IsHeader ? "th" : "td";
            var classes = new List<string>();
            var columnSpan = Math.Max(cell.ColumnSpan, 1);
            if (columnWidths.Count > 0)
            {
                var coveredPercent = 0;
                var countedColumns = 0;
                for (var offset = 0; offset < columnSpan && columnIndex + offset < columnWidths.Count; offset++)
                {
                    coveredPercent += columnWidths[columnIndex + offset];
                    countedColumns++;
                }

                if (countedColumns > 0)
                {
                    classes.Add(registry.Register(
                        "w-" + coveredPercent.ToString(CultureInfo.InvariantCulture),
                        "width: " + coveredPercent.ToString(CultureInfo.InvariantCulture) + "%;"));
                }
            }

            if (cell.RowSpan > 1)
            {
                for (var offset = 0; offset < columnSpan && columnIndex + offset < rowsRemainingInSpan.Length; offset++)
                    rowsRemainingInSpan[columnIndex + offset] = cell.RowSpan;
            }

            columnIndex += columnSpan;
            if (!string.IsNullOrEmpty(cell.ShadingHex))
            {
                var hex = cell.ShadingHex.ToLowerInvariant();
                classes.Add(registry.Register("bg-" + hex, "background-color: #" + hex + ";"));
            }

            if (!string.IsNullOrEmpty(cell.HorizontalAlignment))
                classes.Add(registry.Register("align-" + cell.HorizontalAlignment, "text-align: " + cell.HorizontalAlignment + ";"));

            if (!string.IsNullOrEmpty(cell.VerticalAlignment))
                classes.Add(registry.Register("valign-" + cell.VerticalAlignment, "vertical-align: " + cell.VerticalAlignment + ";"));

            var unwrapParagraph = cell.Blocks.Count == 1
                && cell.Blocks[0] is ParagraphBlock onlyParagraph
                && !IsHeadingTag(TagFor(onlyParagraph));
            if (unwrapParagraph && cell.Blocks[0] is ParagraphBlock paragraph)
                classes.AddRange(ClassesFor(paragraph));

            html.Append('<').Append(tag);
            if (cell.ColumnSpan > 1)
                html.Append(" colspan=\"").Append(cell.ColumnSpan.ToString(CultureInfo.InvariantCulture)).Append('"');
            if (cell.RowSpan > 1)
                html.Append(" rowspan=\"").Append(cell.RowSpan.ToString(CultureInfo.InvariantCulture)).Append('"');
            AppendClass(classes);
            html.Append('>');
            if (unwrapParagraph && cell.Blocks[0] is ParagraphBlock unwrapped)
                WriteInlines(unwrapped.Inlines, unwrapped.Baseline, allowTextBox: true);
            else
                WriteBlocks(cell.Blocks);
            html.Append("</").Append(tag).Append(">\n");
        }

        for (var index = 0; index < rowsRemainingInSpan.Length; index++)
        {
            if (rowsRemainingInSpan[index] > 0)
                rowsRemainingInSpan[index]--;
        }

        html.Append("</tr>\n");
    }

    private void WriteTextBox(IReadOnlyList<Block> blocks)
    {
        html.Append("<div class=\"").Append(FeatureClass(
            "textbox",
            "border: 1px solid #cccccc;",
            "padding: 0.75em;",
            "margin: 0.5em 0;",
            "page-break-inside: avoid;")).Append("\">\n");
        WriteBlocks(blocks);
        html.Append("</div>\n");
    }

    private void WriteNotes(bool endnotes)
    {
        var relevant = document.Notes.Where(note => note.IsEndnote == endnotes).OrderBy(note => note.DisplayNumber).ToList();
        if (relevant.Count == 0)
            return;

        var sectionClass = FeatureClass(
            endnotes ? "endnotes" : "footnotes",
            "margin-top: 2em;",
            "border-top: 1px solid #cccccc;",
            "font-size: 0.92em;");
        var marker = endnotes ? "en" : "fn";
        html.Append("<section class=\"").Append(sectionClass).Append("\">\n<ol>\n");
        foreach (var note in relevant)
        {
            html.Append("<li id=\"").Append(marker).Append('-').Append(note.DisplayNumber.ToString(CultureInfo.InvariantCulture)).Append("\">");
            WriteNoteBody(note.Blocks);
            html.Append(" <a href=\"#").Append(marker).Append("ref-").Append(note.DisplayNumber.ToString(CultureInfo.InvariantCulture))
                .Append("\" class=\"").Append(Prefix("footnote-backlink")).Append("\">&#8617;</a>");
            html.Append("</li>\n");
        }
        html.Append("</ol>\n</section>\n");
    }

    private void WriteNoteBody(IReadOnlyList<Block> blocks)
    {
        if (blocks.Count == 1 && blocks[0] is ParagraphBlock paragraph && !IsHeadingTag(TagFor(paragraph)))
        {
            WriteInlines(paragraph.Inlines, paragraph.Baseline, allowTextBox: true);
            return;
        }

        WriteBlocks(blocks);
    }

    private void WriteInlines(IReadOnlyList<Inline> inlines, RunFormatting baseline, bool allowTextBox)
    {
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case TextRunInline textRun:
                    WriteTextRun(textRun, baseline);
                    break;
                case LineBreakInline:
                    html.Append("<br>");
                    break;
                case PageBreakInline:
                    html.Append("<hr class=\"").Append(FeatureClass(
                        "page-break",
                        "page-break-before: always;",
                        "border: 0;",
                        "border-top: 1px solid #cccccc;",
                        "margin: 1.5em 0;",
                        "height: 0;")).Append("\">");
                    break;
                case TabInline:
                    html.Append("<span class=\"").Append(FeatureClass("tab", "display: inline-block;", "width: 2em;")).Append("\"></span>");
                    break;
                case NonBreakingHyphenInline:
                    html.Append("&#8209;");
                    break;
                case SoftHyphenInline:
                    html.Append("&#173;");
                    break;
                case SymbolInline symbol:
                    html.Append(HtmlText.EncodePreservingSpaces(symbol.Character));
                    break;
                case BookmarkInline bookmark:
                    html.Append("<a id=\"").Append(MarkupSafety.Encode(bookmark.Identifier)).Append("\"></a>");
                    break;
                case HyperlinkInline hyperlink:
                    html.Append("<a href=\"").Append(MarkupSafety.Encode(hyperlink.Href)).Append("\">");
                    WriteInlines(hyperlink.Children, baseline, allowTextBox);
                    html.Append("</a>");
                    break;
                case ImageInline image:
                    WriteImage(image);
                    break;
                case NoteReferenceInline note:
                    WriteNoteReference(note);
                    break;
                case CommentInline comment:
                    html.Append("<span class=\"").Append(FeatureClass("comment", "color: #666666;", "font-size: 0.9em;")).Append("\">")
                        .Append(HtmlText.EncodePreservingSpaces(comment.Text))
                        .Append("</span>");
                    break;
                case TextBoxInline textBox when allowTextBox:
                    WriteTextBox(textBox.Blocks);
                    break;
            }
        }
    }

    private void WriteTextRun(TextRunInline textRun, RunFormatting baseline)
    {
        var effective = textRun.Formatting;
        var className = RunClass(effective, baseline);
        var bold = options.DirectFormatting != DirectFormattingMode.None && effective.Bold && !baseline.Bold;
        var italic = options.DirectFormatting != DirectFormattingMode.None && effective.Italic && !baseline.Italic;
        var underline = options.DirectFormatting != DirectFormattingMode.None && effective.Underline && !baseline.Underline;
        var strike = options.DirectFormatting != DirectFormattingMode.None && effective.Strike && !baseline.Strike;
        var superscript = options.DirectFormatting != DirectFormattingMode.None
            && effective.VerticalPosition == VerticalTextPosition.Superscript
            && baseline.VerticalPosition != VerticalTextPosition.Superscript;
        var subscript = options.DirectFormatting != DirectFormattingMode.None
            && effective.VerticalPosition == VerticalTextPosition.Subscript
            && baseline.VerticalPosition != VerticalTextPosition.Subscript;

        if (className is not null)
            html.Append("<span class=\"").Append(className).Append("\">");
        if (superscript)
            html.Append("<sup>");
        else if (subscript)
            html.Append("<sub>");
        if (strike)
            html.Append("<s>");
        if (underline)
            html.Append("<u>");
        if (italic)
            html.Append("<em>");
        if (bold)
            html.Append("<strong>");

        html.Append(HtmlText.EncodePreservingSpaces(textRun.Text));

        if (bold)
            html.Append("</strong>");
        if (italic)
            html.Append("</em>");
        if (underline)
            html.Append("</u>");
        if (strike)
            html.Append("</s>");
        if (superscript)
            html.Append("</sup>");
        else if (subscript)
            html.Append("</sub>");
        if (className is not null)
            html.Append("</span>");
    }

    private void WriteImage(ImageInline image)
    {
        if (!image.SourceIsUri && image.Data.Length == 0)
            return;

        string source;
        if (image.SourceIsUri)
            source = image.FileName;
        else if (options.ImageMode == ImageMode.DataUri)
            source = "data:" + image.ContentType + ";base64," + Convert.ToBase64String(image.Data);
        else
        {
            var folder = HtmlDocumentWriter.ImageFolder(options);
            source = string.IsNullOrEmpty(folder) ? image.FileName : folder + "/" + image.FileName;
        }

        string? className = null;
        if (image.WidthPixels is int width && width > 0)
        {
            if (image.HeightPixels is int height && height > 0)
            {
                className = registry.Register(
                    "img-w-" + width.ToString(CultureInfo.InvariantCulture) + "-h-" + height.ToString(CultureInfo.InvariantCulture),
                    "width: " + width.ToString(CultureInfo.InvariantCulture) + "px;",
                    "height: " + height.ToString(CultureInfo.InvariantCulture) + "px;");
            }
            else
            {
                className = registry.Register(
                    "img-w-" + width.ToString(CultureInfo.InvariantCulture),
                    "width: " + width.ToString(CultureInfo.InvariantCulture) + "px;",
                    "height: auto;");
            }
        }

        html.Append("<img");
        if (className is not null)
            html.Append(" class=\"").Append(className).Append('"');
        html.Append(" src=\"").Append(MarkupSafety.Encode(source))
            .Append("\" alt=\"").Append(MarkupSafety.Encode(image.AltText)).Append("\">");
    }

    private void WriteNoteReference(NoteReferenceInline note)
    {
        var marker = note.IsEndnote ? "en" : "fn";
        var number = note.DisplayNumber.ToString(CultureInfo.InvariantCulture);
        html.Append("<sup><a href=\"#").Append(marker).Append('-').Append(number)
            .Append("\" id=\"").Append(marker).Append("ref-").Append(number)
            .Append("\" class=\"").Append(FeatureClass("footnote-ref", "text-decoration: none;")).Append("\">")
            .Append(number)
            .Append("</a></sup>");
    }

    private string? RunClass(RunFormatting effective, RunFormatting baseline)
    {
        if (options.DirectFormatting != DirectFormattingMode.SemanticAndClasses)
            return null;

        var tokens = new List<string>();
        var declarations = new List<string>();
        void Add(string token, string declaration)
        {
            tokens.Add(token);
            declarations.Add(declaration);
        }

        if (!effective.Bold && baseline.Bold)
            Add("weight-normal", "font-weight: normal;");
        if (!effective.Italic && baseline.Italic)
            Add("style-normal", "font-style: normal;");
        if (!effective.Underline && baseline.Underline)
            Add("no-underline", "text-decoration: none;");
        if (!effective.Strike && baseline.Strike)
            Add("no-strike", "text-decoration: none;");

        if (!string.Equals(effective.ColorHex, baseline.ColorHex, StringComparison.OrdinalIgnoreCase))
        {
            if (effective.ColorHex is null)
                Add("color-inherit", "color: inherit;");
            else
                Add("color-" + effective.ColorHex.ToLowerInvariant(), "color: #" + effective.ColorHex.ToLowerInvariant() + ";");
        }

        if (!string.Equals(effective.HighlightHex, baseline.HighlightHex, StringComparison.OrdinalIgnoreCase))
        {
            if (effective.HighlightHex is null)
                Add("highlight-transparent", "background-color: transparent;");
            else
                Add("highlight-" + effective.HighlightHex.ToLowerInvariant(), "background-color: #" + effective.HighlightHex.ToLowerInvariant() + ";");
        }

        if (effective.FontSizeHalfPoints != baseline.FontSizeHalfPoints && effective.FontSizeHalfPoints is int halfPoints)
        {
            var points = UnitConverter.HalfPointsToPoints(halfPoints);
            Add("size-" + CssNumbers.PointToken(points) + "pt", "font-size: " + CssNumbers.Points(points) + "pt;");
        }

        if (!string.Equals(effective.FontFamily, baseline.FontFamily, StringComparison.OrdinalIgnoreCase)
            && !string.IsNullOrWhiteSpace(effective.FontFamily))
        {
            Add("font-" + StyleSlug.Create(effective.FontFamily), "font-family: " + CssBuilder.QuoteFont(effective.FontFamily) + ";");
        }

        if (tokens.Count == 0)
            return null;

        return registry.Register(string.Join('-', tokens), declarations.ToArray());
    }

    private void AddDirectParagraphClasses(ParagraphBlock paragraph, List<string> classes)
    {
        if (options.DirectFormatting != DirectFormattingMode.SemanticAndClasses)
            return;

        var style = document.Styles.Paragraph(paragraph.StyleId);
        var direct = paragraph.DirectFormatting;
        if (direct.Alignment is string alignment && !string.Equals(alignment, style.Alignment, StringComparison.Ordinal))
            classes.Add(registry.Register("align-" + alignment, "text-align: " + alignment + ";"));

        var tokens = new List<string>();
        var declarations = new List<string>();
        AddSpacingClass(tokens, declarations, "mt", direct.SpacingBeforeTwips, style.SpacingBeforeTwips, "margin-top");
        AddSpacingClass(tokens, declarations, "mb", direct.SpacingAfterTwips, style.SpacingAfterTwips, "margin-bottom");
        AddSpacingClass(tokens, declarations, "ml", direct.IndentLeftTwips, style.IndentLeftTwips, "margin-left");
        AddSpacingClass(tokens, declarations, "mr", direct.IndentRightTwips, style.IndentRightTwips, "margin-right");
        if (direct.FirstLineTwips is int firstLine)
            AddSpacingClass(tokens, declarations, "ti", firstLine, style.FirstLineTwips, "text-indent");
        else if (direct.HangingTwips is int hanging)
            AddSpacingClass(tokens, declarations, "ti-n", hanging, style.HangingTwips, "text-indent", negative: true);

        if (tokens.Count > 0)
            classes.Add(registry.Register(string.Join('-', tokens), declarations.ToArray()));
    }

    private static void AddSpacingClass(
        List<string> tokens,
        List<string> declarations,
        string name,
        int? direct,
        int? styleValue,
        string property,
        bool negative = false)
    {
        if (direct is not int twips)
            return;

        var snapped = SnapSpacingPoints(UnitConverter.TwipsToPoints(twips));
        var styleSnapped = styleValue is int styleTwips ? SnapSpacingPoints(UnitConverter.TwipsToPoints(styleTwips)) : 0;
        if (snapped == styleSnapped)
            return;

        tokens.Add(name + "-" + snapped.ToString(CultureInfo.InvariantCulture));
        var points = (negative ? "-" : "") + snapped.ToString(CultureInfo.InvariantCulture) + "pt";
        declarations.Add(property + ": " + points + ";");
    }

    private static int SnapSpacingPoints(double points)
    {
        var snapped = Math.Round(points / 4d, MidpointRounding.AwayFromZero) * 4d;
        if (snapped < 0)
            snapped = 0;
        if (snapped > 96)
            snapped = 96;

        return (int)snapped;
    }

    private List<string> ClassesFor(ParagraphBlock paragraph)
    {
        var classes = new List<string>();
        switch (paragraph.Kind)
        {
            case ParagraphKind.Title:
                classes.Add(Prefix("title"));
                Remember(paragraph.StyleId, Prefix("title"), isTable: false);
                break;
            case ParagraphKind.Subtitle:
                classes.Add(Prefix("subtitle"));
                Remember(paragraph.StyleId, Prefix("subtitle"), isTable: false);
                break;
            case ParagraphKind.Mapped:
                if (!string.IsNullOrWhiteSpace(paragraph.MappedClass))
                {
                    var first = true;
                    foreach (var token in paragraph.MappedClass.Split(' ', StringSplitOptions.RemoveEmptyEntries))
                    {
                        var className = MarkupSafety.ClassToken(token);
                        classes.Add(className);
                        if (first)
                        {
                            Remember(paragraph.StyleId, className, isTable: false);
                            first = false;
                        }
                    }
                }

                if (HeadingMapper.LevelFromTag(paragraph.MappedTag) is int mappedLevel and > 6)
                    classes.Add(Prefix("h" + mappedLevel.ToString(CultureInfo.InvariantCulture)));
                break;
            default:
                if (!paragraph.IsNormalStyle && !string.IsNullOrWhiteSpace(paragraph.StyleName))
                {
                    var className = Prefix("p-" + StyleSlug.Create(paragraph.StyleName));
                    classes.Add(className);
                    Remember(paragraph.StyleId, className, isTable: false);
                }

                if (paragraph.HeadingLevel is int level && level > 6)
                    classes.Add(Prefix("h" + level.ToString(CultureInfo.InvariantCulture)));
                break;
        }

        AddDirectParagraphClasses(paragraph, classes);
        return classes;
    }

    private string TagFor(ParagraphBlock paragraph)
    {
        if (paragraph.Kind == ParagraphKind.Mapped)
        {
            var mapped = SanitizeTag(paragraph.MappedTag);
            if (HeadingMapper.LevelFromTag(paragraph.MappedTag) is > 6)
                return "h6";
            return mapped;
        }

        if (paragraph.Kind == ParagraphKind.Title && options.UseSemanticHeadings)
            return "h1";

        if (paragraph.HeadingLevel is int level && options.UseSemanticHeadings)
            return "h" + Math.Min(level, 6).ToString(CultureInfo.InvariantCulture);

        return "p";
    }

    private void CollectHeading(ParagraphBlock paragraph, string tag, IReadOnlyList<Inline> inlines)
    {
        if (!IsHeadingTag(tag))
            return;

        int? level = paragraph.HeadingLevel;
        if (paragraph.Kind == ParagraphKind.Title)
            level ??= 1;
        if (paragraph.Kind == ParagraphKind.Mapped)
            level = HeadingMapper.LevelFromTag(paragraph.MappedTag) ?? level;

        if (level is null or < 1)
            return;

        headings.Add(new CollectedHeading(level.Value, PlainText(inlines), paragraph.ElementId));
    }

    private static bool IsHeadingTag(string tag) => tag is "h1" or "h2" or "h3" or "h4" or "h5" or "h6";

    private static string SanitizeTag(string? tag)
    {
        if (string.IsNullOrWhiteSpace(tag))
            return "p";

        var builder = new StringBuilder();
        foreach (var character in tag.Trim().ToLowerInvariant())
        {
            if (char.IsAsciiLetterOrDigit(character))
                builder.Append(character);
        }

        var sanitized = builder.ToString();
        return sanitized switch
        {
            "h1" or "h2" or "h3" or "h4" or "h5" or "h6" or "p" or "blockquote" or "pre"
                or "div" or "section" or "article" or "aside" or "figure" or "figcaption"
                or "address" => sanitized,
            _ when HeadingMapper.LevelFromTag(sanitized) is > 6 => "h6",
            _ => "p"
        };
    }

    private void Remember(string? styleId, string className, bool isTable)
    {
        if (string.IsNullOrEmpty(styleId) || string.IsNullOrEmpty(className))
            return;

        var key = (isTable ? "t:" : "p:") + className;
        if (!rememberedStyles.Add(key))
            return;

        usedStyles.Add(new UsedStyle
        {
            ClassName = className,
            StyleId = styleId,
            IsTable = isTable
        });
    }

    private string Prefix(string className) => MarkupSafety.ClassToken((options.ClassPrefix ?? "") + className);

    private string FeatureClass(string name, params string[] declarations)
    {
        if (options.IncludeBaseCss)
            registry.Register(name, declarations);

        return Prefix(name);
    }

    private void AppendClass(IReadOnlyList<string> classes)
    {
        var distinct = classes.Where(className => !string.IsNullOrEmpty(className)).Distinct(StringComparer.Ordinal).ToList();
        if (distinct.Count == 0)
            return;

        html.Append(" class=\"").Append(string.Join(' ', distinct)).Append('"');
    }

    private static string PlainText(IEnumerable<Inline> inlines)
    {
        var builder = new StringBuilder();
        AppendPlain(inlines, builder);
        return HtmlText.CollapseWhitespace(builder.ToString());
    }

    private static void AppendPlain(IEnumerable<Inline> inlines, StringBuilder builder)
    {
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case TextRunInline textRun:
                    builder.Append(textRun.Text);
                    break;
                case SymbolInline symbol:
                    builder.Append(symbol.Character);
                    break;
                case HyperlinkInline hyperlink:
                    AppendPlain(hyperlink.Children, builder);
                    break;
                case TabInline:
                case LineBreakInline:
                    builder.Append(' ');
                    break;
            }
        }
    }

    private static List<FlowSegment> SplitTextBoxes(IReadOnlyList<Inline> inlines)
    {
        var segments = new List<FlowSegment>();
        var current = new List<Inline>();
        foreach (var inline in inlines)
        {
            if (inline is TextBoxInline textBox)
            {
                if (current.Count > 0)
                {
                    segments.Add(new FlowSegment(current, null));
                    current = new List<Inline>();
                }

                segments.Add(new FlowSegment(new List<Inline>(), textBox));
            }
            else
            {
                current.Add(inline);
            }
        }

        if (current.Count > 0 || segments.Count == 0)
            segments.Add(new FlowSegment(current, null));

        return segments;
    }

    private sealed record FlowSegment(List<Inline> Inlines, TextBoxInline? TextBox);
}
