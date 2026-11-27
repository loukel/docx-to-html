using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocxToHtml.Tests.Support;

internal static class WordFixtureFactory
{
    public static readonly string[] Names =
    [
        "headings",
        "paragraph-styles",
        "run-formatting",
        "fragmented-runs",
        "hyperlinks",
        "fields",
        "tracked-changes",
        "lists",
        "simple-table",
        "header-row",
        "colspan",
        "rowspan",
        "colspan-rowspan",
        "nested-table",
        "table-in-content-control",
        "images",
        "text-box",
        "footnotes"
    ];

    public static byte[] Build(string name) => name switch
    {
        "headings" => BuildHeadings(),
        "paragraph-styles" => BuildParagraphStyles(),
        "run-formatting" => BuildRunFormatting(),
        "fragmented-runs" => BuildFragmentedRuns(),
        "hyperlinks" => BuildHyperlinks(),
        "fields" => BuildFields(),
        "tracked-changes" => BuildTrackedChanges(),
        "lists" => BuildLists(),
        "simple-table" => BuildSimpleTable(),
        "header-row" => BuildHeaderRow(),
        "colspan" => BuildColumnSpan(),
        "rowspan" => BuildRowSpan(),
        "colspan-rowspan" => BuildColumnAndRowSpan(),
        "nested-table" => BuildNestedTable(),
        "table-in-content-control" => BuildTableInContentControl(),
        "images" => BuildImages(),
        "text-box" => BuildTextBox(),
        "footnotes" => BuildFootnotes(),
        _ => throw new ArgumentException($"Unknown fixture '{name}'.", nameof(name))
    };

    private static byte[] BuildHeadings()
    {
        var builder = new WordDocumentBuilder("Headings");
        for (var level = 1; level <= 9; level++)
            builder.AddParagraphStyle($"Heading{level}", $"heading {level}", outlineLevel: level - 1, bold: true, halfPoints: 32 - Math.Min(level, 6) * 2);
        builder.AddParagraphStyle("Title", "Title", halfPoints: 40, bold: true);
        builder.AddParagraphStyle("Subtitle", "Subtitle", halfPoints: 24, italic: true, color: "595959");
        builder.AddParagraphStyle("AppendixHeading", "Appendix Heading", basedOn: "Heading2", bold: true, halfPoints: 28);
        builder.AddParagraphStyle("DeepCustom", "Deep Custom", basedOn: "Heading7", bold: true, halfPoints: 22);
        builder.AddParagraph("Chapter", "Heading1", bold: true);
        builder.AddParagraph("Section", "Heading2");
        builder.AddParagraph("Topic", "Heading3");
        builder.AddParagraph("Detail", "Heading4");
        builder.AddParagraph("Note", "Heading5");
        builder.AddParagraph("Aside", "Heading6");
        builder.AddParagraph("Deep", "Heading7");
        builder.AddParagraph("Deeper", "Heading8");
        builder.AddParagraph("Deepest", "Heading9");
        builder.AddParagraph("Document title", "Title");
        builder.AddParagraph("A subtitle", "Subtitle");
        builder.AddParagraph("Appendix", "AppendixHeading");
        builder.AddParagraph("Buried", "DeepCustom");
        builder.AddParagraph("Outlined", outlineLevel: 0);
        builder.AddParagraph("Body text");
        return builder.Finish();
    }

    private static byte[] BuildParagraphStyles()
    {
        var builder = new WordDocumentBuilder("Paragraph styles");
        builder.AddParagraphStyle("Quote", "Quote", italic: true, color: "1F4E79", indentLeft: 720, spacingAfter: 120);
        builder.AddParagraphStyle("IntenseQuote", "Intense Quote", basedOn: "Quote", bold: true);
        builder.AddParagraph("Just a sentence.");
        builder.AddParagraph("A quoted line.", "Quote");
        builder.AddParagraph("An intense line.", "IntenseQuote");
        builder.Add(new Paragraph());
        return builder.Finish();
    }

    private static byte[] BuildRunFormatting()
    {
        var builder = new WordDocumentBuilder("Run formatting");
        builder.AddParagraph("Plain words");
        builder.AddParagraph(Run("Bold words", bold: true));
        builder.AddParagraph(Run("Italic words", italic: true));
        builder.AddParagraph(Run("Underlined words", underline: true));
        builder.AddParagraph(Run("Struck words", strike: true));
        builder.AddParagraph(Run("Superscript", vertical: "superscript"));
        builder.AddParagraph(Run("Subscript", vertical: "subscript"));
        builder.AddParagraph(Run("Red words", color: "FF0000"));
        builder.AddParagraph(Run("Highlighted words", highlight: "yellow"));
        builder.AddParagraph(Run("Bigger words", halfPoints: "28"));
        builder.AddAlignedParagraph(Run("Centered words"), "center");
        return builder.Finish();
    }

    private static byte[] BuildFragmentedRuns()
    {
        var builder = new WordDocumentBuilder("Fragmented runs");
        var paragraph = new Paragraph();
        paragraph.Append(new Run(
            new RunProperties(new Bold(), new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" }),
            new Text("Hel")));
        paragraph.Append(new ProofError { Type = ProofingErrorValues.SpellStart });
        paragraph.Append(new Run(
            new RunProperties(
                new Bold(),
                new Languages { Val = "en-US" },
                new NoProof(),
                new FontSizeComplexScript { Val = "22" }),
            new Text("lo") { Space = SpaceProcessingModeValues.Preserve }));
        paragraph.Append(new ProofError { Type = ProofingErrorValues.SpellEnd });
        builder.Add(paragraph);
        return builder.Finish();
    }

    private static byte[] BuildHyperlinks()
    {
        var builder = new WordDocumentBuilder("Hyperlinks");
        builder.AddParagraph(builder.ExternalLink("Example", "https://example.com/"));
        builder.AddParagraph(
            new BookmarkStart { Id = "1", Name = "chapter" },
            new Run(new Text("Chapter")),
            new BookmarkEnd { Id = "1" });
        builder.AddParagraph(new Hyperlink(new Run(new Text("Go to chapter"))) { Anchor = "chapter" });
        return builder.Finish();
    }

    private static byte[] BuildFields()
    {
        var builder = new WordDocumentBuilder("Fields");
        builder.AddParagraph(
            FieldCharRun(FieldCharValues.Begin),
            new Run(new FieldCode(" HYPERLINK \"https://example.com/complex\" ") { Space = SpaceProcessingModeValues.Preserve }),
            FieldCharRun(FieldCharValues.Separate),
            new Run(new Text("Complex link")),
            FieldCharRun(FieldCharValues.End),
            new Run(new Text(" and ") { Space = SpaceProcessingModeValues.Preserve }),
            new SimpleField(new Run(new Text("Simple link"))) { Instruction = " HYPERLINK \"https://example.com/simple\" " },
            new Run(new Text(" on ") { Space = SpaceProcessingModeValues.Preserve }),
            FieldCharRun(FieldCharValues.Begin),
            new Run(new FieldCode(" DATE ") { Space = SpaceProcessingModeValues.Preserve }),
            FieldCharRun(FieldCharValues.Separate),
            new Run(new Text("27 November 2026")),
            FieldCharRun(FieldCharValues.End));
        return builder.Finish();
    }

    private static byte[] BuildTrackedChanges()
    {
        var builder = new WordDocumentBuilder("Tracked changes");
        builder.AddParagraph(
            new Run(new Text("Keep ") { Space = SpaceProcessingModeValues.Preserve }),
            new InsertedRun(new Run(new Text("added"))) { Author = "Ada", Id = "1" },
            new DeletedRun(new Run(new DeletedText("removed") { Space = SpaceProcessingModeValues.Preserve })) { Author = "Ada", Id = "2" });
        return builder.Finish();
    }

    private static byte[] BuildLists()
    {
        var builder = new WordDocumentBuilder("Lists");
        builder.AddNumbering();
        builder.AddListItem("Fruit", 1, 0);
        builder.AddListItem("Apple", 1, 1);
        builder.AddListItem("Pear", 1, 1);
        builder.AddListItem("Vegetable", 1, 0);
        builder.AddParagraph("Between the lists.");
        builder.AddListItem("First", 2, 0);
        builder.AddListItem("Second", 2, 0);
        builder.AddListItem("Letter", 2, 1);
        builder.AddListItem("Fifth", 3, 0);
        return builder.Finish();
    }

    private static byte[] BuildSimpleTable()
    {
        var builder = new WordDocumentBuilder("Simple table");
        builder.AddTableStyle();
        builder.Add(Table(
            Row(ShadedCell("Name", "D9E2F3"), Cell("Ada")),
            Row(Cell("Role"), CenteredCell("Engineer"))));
        return builder.Finish();
    }

    private static byte[] BuildHeaderRow()
    {
        var builder = new WordDocumentBuilder("Header row");
        builder.AddTableStyle();
        builder.Add(Table(
            Row(true, Cell("Name"), Cell("Role")),
            Row(Cell("Ada"), Cell("Engineer"))));
        return builder.Finish();
    }

    private static byte[] BuildColumnSpan()
    {
        var builder = new WordDocumentBuilder("Column span");
        builder.AddTableStyle();
        builder.Add(Table(
            Row(Cell("Title", columnSpan: 2)),
            Row(Cell("Left"), Cell("Right"))));
        return builder.Finish();
    }

    private static byte[] BuildRowSpan()
    {
        var builder = new WordDocumentBuilder("Row span");
        builder.AddTableStyle();
        builder.Add(Table(
            Row(Cell("Ada", verticalMerge: MergedCellValues.Restart), Cell("Engineer")),
            Row(Cell("", verticalMerge: MergedCellValues.Continue), Cell("Writer"))));
        return builder.Finish();
    }

    private static byte[] BuildColumnAndRowSpan()
    {
        var builder = new WordDocumentBuilder("Column and row span");
        builder.AddTableStyle();
        builder.Add(Table(
            Row(Cell("Ada", columnSpan: 2, verticalMerge: MergedCellValues.Restart), Cell("Engineer")),
            Row(Cell("", columnSpan: 2, verticalMerge: MergedCellValues.Continue), Cell("Writer"))));
        return builder.Finish();
    }

    private static byte[] BuildNestedTable()
    {
        var builder = new WordDocumentBuilder("Nested table");
        builder.AddTableStyle();
        var inner = Table(Row(Cell("Inner")));
        var outerCell = new TableCell(new Paragraph(new Run(new Text("Outer"))), inner);
        builder.Add(Table(new TableRow(outerCell)));
        return builder.Finish();
    }

    private static byte[] BuildTableInContentControl()
    {
        var builder = new WordDocumentBuilder("Content control");
        builder.AddTableStyle();
        builder.Add(new SdtBlock(new SdtProperties(new SdtId { Val = 1 }), new SdtContentBlock(Table(Row(Cell("Inside"))))));
        return builder.Finish();
    }

    private static byte[] BuildImages()
    {
        var builder = new WordDocumentBuilder("Images");
        builder.AddImageParagraph("Sample chart", 3048000, 1524000);
        return builder.Finish();
    }

    private static byte[] BuildTextBox()
    {
        var builder = new WordDocumentBuilder("Text box");
        var paragraph = new Paragraph(
            new Run(new Text("Label ") { Space = SpaceProcessingModeValues.Preserve }));
        var drawingRun = new Run();
        drawingRun.InnerXml = """
            <w:pict xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml">
              <v:shape style="width:120pt;height:40pt">
                <v:textbox>
                  <w:txbxContent>
                    <w:p><w:r><w:t>Inside the box</w:t></w:r></w:p>
                  </w:txbxContent>
                </v:textbox>
              </v:shape>
            </w:pict>
            """;
        paragraph.Append(drawingRun);
        builder.Add(paragraph);
        return builder.Finish();
    }

    private static byte[] BuildFootnotes()
    {
        var builder = new WordDocumentBuilder("Footnotes");
        builder.AddHeaderAndFooter("Running header", "Running footer");
        var reference = new Run(new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }), new FootnoteReference { Id = 1 });
        builder.AddParagraph(new Run(new Text("See source")), reference);
        builder.AddFootnote(1, "A note about sources");
        return builder.Finish();
    }

    private static Run Run(string text, bool bold = false, bool italic = false, bool underline = false, bool strike = false, string? vertical = null, string? color = null, string? highlight = null, string? halfPoints = null)
    {
        var properties = new RunProperties();
        if (bold)
            properties.Append(new Bold());
        if (italic)
            properties.Append(new Italic());
        if (underline)
            properties.Append(new Underline { Val = UnderlineValues.Single });
        if (strike)
            properties.Append(new Strike());
        if (vertical == "superscript")
            properties.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript });
        if (vertical == "subscript")
            properties.Append(new VerticalTextAlignment { Val = VerticalPositionValues.Subscript });
        if (color is not null)
            properties.Append(new Color { Val = color });
        if (highlight is not null)
            properties.Append(new Highlight { Val = HighlightColorValues.Yellow });
        if (halfPoints is not null)
            properties.Append(new FontSize { Val = halfPoints });

        return properties.HasChildren
            ? new Run(properties, new Text(text) { Space = SpaceProcessingModeValues.Preserve })
            : new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
    }

    private static Run FieldCharRun(FieldCharValues type) => new(new FieldChar { FieldCharType = type });

    private static Table Table(params OpenXmlElement[] rows)
    {
        var table = new Table(
            new TableProperties(
                new TableStyle { Val = "TableGrid" },
                new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                    new LeftBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                    new BottomBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                    new RightBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Space = 0, Color = "auto" })),
            new TableGrid(new GridColumn { Width = "2500" }, new GridColumn { Width = "2500" }));
        foreach (var row in rows)
            table.Append(row);
        return table;
    }

    private static TableRow Row(params TableCell[] cells) => Row(false, cells);

    private static TableRow Row(bool header, params TableCell[] cells)
    {
        var row = new TableRow();
        if (header)
            row.Append(new TableRowProperties(new TableHeader()));
        foreach (var cell in cells)
            row.Append(cell);
        return row;
    }

    private static TableCell Cell(string text, int? columnSpan = null, MergedCellValues? verticalMerge = null)
    {
        var cell = new TableCell();
        var properties = new TableCellProperties();
        if (columnSpan is int span)
            properties.Append(new GridSpan { Val = span });
        if (verticalMerge is MergedCellValues merge)
            properties.Append(new VerticalMerge { Val = merge });
        if (properties.HasChildren)
            cell.Append(properties);
        cell.Append(string.IsNullOrEmpty(text) ? new Paragraph() : new Paragraph(new Run(new Text(text))));
        return cell;
    }

    private static TableCell ShadedCell(string text, string fill)
    {
        var cell = new TableCell(
            new TableCellProperties(new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = fill }),
            new Paragraph(new Run(new Text(text))));
        return cell;
    }

    private static TableCell CenteredCell(string text) =>
        new(new Paragraph(
            new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
            new Run(new Text(text))));
}

internal sealed class WordDocumentBuilder
{
    private readonly string path = Path.Combine(Path.GetTempPath(), $"docxtohtml-{Guid.NewGuid():N}.docx");
    private readonly WordprocessingDocument document;
    private readonly MainDocumentPart main;
    private readonly Body body;
    private Styles? styles;
    private bool hasSection;

    public WordDocumentBuilder(string title)
    {
        document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        main = document.AddMainDocumentPart();
        body = new Body();
        main.Document = new Document(body);
        document.PackageProperties.Title = title;
        EnsureStyles();
    }

    public void AddParagraphStyle(string styleId, string name, string? basedOn = null, int? outlineLevel = null, bool bold = false, bool italic = false, string? color = null, int? halfPoints = null, int? indentLeft = null, int? spacingAfter = null)
    {
        var style = new Style { Type = StyleValues.Paragraph, StyleId = styleId };
        style.Append(new StyleName { Val = name });
        if (basedOn is not null)
        {
            style.CustomStyle = true;
            style.Append(new BasedOn { Val = basedOn });
        }

        var paragraphProperties = new StyleParagraphProperties();
        if (outlineLevel is int outline)
            paragraphProperties.Append(new OutlineLevel { Val = outline });
        if (indentLeft is int left)
            paragraphProperties.Append(new Indentation { Left = left.ToString(System.Globalization.CultureInfo.InvariantCulture) });
        if (spacingAfter is int after)
            paragraphProperties.Append(new SpacingBetweenLines { After = after.ToString(System.Globalization.CultureInfo.InvariantCulture) });
        if (paragraphProperties.HasChildren)
            style.Append(paragraphProperties);

        var runProperties = new StyleRunProperties();
        if (bold)
            runProperties.Append(new Bold());
        if (italic)
            runProperties.Append(new Italic());
        if (color is not null)
            runProperties.Append(new Color { Val = color });
        if (halfPoints is int size)
            runProperties.Append(new FontSize { Val = size.ToString(System.Globalization.CultureInfo.InvariantCulture) });
        if (runProperties.HasChildren)
            style.Append(runProperties);

        styles!.Append(style);
    }

    public void AddParagraph(string text, string? styleId = null, int? outlineLevel = null, bool bold = false)
    {
        AddParagraph(bold ? WordRun(text) : new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }), styleId, outlineLevel);
    }

    public void AddParagraph(params OpenXmlElement[] children) => AddParagraph(children, styleId: null, outlineLevel: null);

    public void AddAlignedParagraph(OpenXmlElement content, string alignment) =>
        AddParagraph(content, styleId: null, outlineLevel: null, alignment);

    public void AddListItem(string text, int numberingId, int level)
    {
        var paragraph = new Paragraph(
            new ParagraphProperties(new NumberingProperties(
                new NumberingLevelReference { Val = level },
                new NumberingId { Val = numberingId })),
            new Run(new Text(text)));
        body.Append(paragraph);
    }

    public void Add(OpenXmlElement element) => body.Append(element);

    public Hyperlink ExternalLink(string text, string href)
    {
        var relationship = main.AddHyperlinkRelationship(new Uri(href, UriKind.Absolute), true);
        return new Hyperlink(new Run(new Text(text))) { Id = relationship.Id };
    }

    public void AddMultiLevelHeadingNumbering()
    {
        var numbering = new Numbering();
        var abstractNumbering = new AbstractNum { AbstractNumberId = 10 };
        abstractNumbering.Append(MultiLevel(0, "%1"));
        abstractNumbering.Append(MultiLevel(1, "%1.%2"));
        abstractNumbering.Append(MultiLevel(2, "%1.%2.%3"));
        abstractNumbering.Append(MultiLevel(3, "%1.%2.%3.%4"));
        numbering.Append(abstractNumbering);
        numbering.Append(new NumberingInstance(new AbstractNumId { Val = 10 }) { NumberID = 4 });
        var part = main.AddNewPart<NumberingDefinitionsPart>();
        part.Numbering = numbering;

        for (var level = 1; level <= 4; level++)
        {
            AddParagraphStyle($"Heading{level}", $"heading {level}", outlineLevel: level - 1, bold: true);
            var style = styles!.Elements<Style>().Last();
            var paragraphProperties = style.Elements<StyleParagraphProperties>().FirstOrDefault();
            if (paragraphProperties is null)
            {
                paragraphProperties = new StyleParagraphProperties();
                style.Append(paragraphProperties);
            }

            paragraphProperties.Append(new NumberingProperties(
                new NumberingLevelReference { Val = level - 1 },
                new NumberingId { Val = 4 }));
        }
    }

    private static Level MultiLevel(int index, string levelText) =>
        new(
            new StartNumberingValue { Val = 1 },
            new NumberingFormat { Val = NumberFormatValues.Decimal },
            new LevelText { Val = levelText },
            new LevelSuffix { Val = LevelSuffixValues.Space })
        { LevelIndex = index };

    public void AddNumbering()
    {
        var numbering = new Numbering();
        numbering.Append(Abstract(0, NumberFormatValues.Bullet, "•", "•"));
        numbering.Append(Abstract(1, NumberFormatValues.Decimal, "%1.", "%2.", levelOneFormat: NumberFormatValues.LowerLetter));
        numbering.Append(Abstract(2, NumberFormatValues.Decimal, "%1.", start: 5));
        numbering.Append(new NumberingInstance(new AbstractNumId { Val = 0 }) { NumberID = 1 });
        numbering.Append(new NumberingInstance(new AbstractNumId { Val = 1 }) { NumberID = 2 });
        numbering.Append(new NumberingInstance(new AbstractNumId { Val = 2 }) { NumberID = 3 });
        var part = main.AddNewPart<NumberingDefinitionsPart>();
        part.Numbering = numbering;
    }

    public void AddTableStyle()
    {
        var style = new Style { Type = StyleValues.Table, StyleId = "TableGrid" };
        style.Append(new StyleName { Val = "Table Grid" });
        style.Append(new StyleTableProperties(new TableBorders(
            new TopBorder { Val = BorderValues.Single, Size = 4, Color = "auto" },
            new LeftBorder { Val = BorderValues.Single, Size = 4, Color = "auto" },
            new BottomBorder { Val = BorderValues.Single, Size = 4, Color = "auto" },
            new RightBorder { Val = BorderValues.Single, Size = 4, Color = "auto" },
            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = "auto" },
            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = "auto" })));
        styles!.Append(style);
    }

    public void AddImageParagraph(string description, long widthEmu, long heightEmu)
    {
        var imagePart = main.AddImagePart(ImagePartType.Png);
        using (var imageStream = imagePart.GetStream(FileMode.Create, FileAccess.Write))
            imageStream.Write(TinyPng);
        var relationshipId = main.GetIdOfPart(imagePart);
        var run = new Run();
        run.InnerXml = $"""
            <w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <wp:inline>
                <wp:extent cx="{widthEmu}" cy="{heightEmu}"/>
                <wp:docPr id="1" name="Picture 1" descr="{description}"/>
                <a:graphic>
                  <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                    <pic:pic>
                      <pic:blipFill>
                        <a:blip r:embed="{relationshipId}"/>
                      </pic:blipFill>
                      <pic:spPr>
                        <a:xfrm>
                          <a:off x="0" y="0"/>
                          <a:ext cx="{widthEmu}" cy="{heightEmu}"/>
                        </a:xfrm>
                      </pic:spPr>
                    </pic:pic>
                  </a:graphicData>
                </a:graphic>
              </wp:inline>
            </w:drawing>
            """;
        AddParagraph(new Run(new Text("Chart ") { Space = SpaceProcessingModeValues.Preserve }), run);
    }

    public void AddHeaderAndFooter(string headerText, string footerText)
    {
        var headerPart = main.AddNewPart<HeaderPart>();
        headerPart.Header = new Header(new Paragraph(new Run(new Text(headerText))));
        var footerPart = main.AddNewPart<FooterPart>();
        footerPart.Footer = new Footer(new Paragraph(new Run(new Text(footerText))));
        body.Append(new SectionProperties(
            new HeaderReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(headerPart) },
            new FooterReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(footerPart) }));
        hasSection = true;
    }

    public void AddFootnote(int identifier, string text)
    {
        var part = main.FootnotesPart ?? main.AddNewPart<FootnotesPart>();
        part.Footnotes ??= new Footnotes(
            new Footnote(new Paragraph()) { Type = FootnoteEndnoteValues.Separator, Id = -1 });
        part.Footnotes.Append(new Footnote(
            new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "FootnoteText" }),
                new Run(new RunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }), new FootnoteReferenceMark()),
                new Run(new Text(" " + text) { Space = SpaceProcessingModeValues.Preserve })))
        { Id = identifier });
    }

    public byte[] Finish()
    {
        if (!hasSection)
            body.Append(new SectionProperties());
        main.Document?.Save();
        document.Dispose();
        var bytes = File.ReadAllBytes(path);
        File.Delete(path);
        return bytes;
    }

    private void AddParagraph(OpenXmlElement first, string? styleId, int? outlineLevel, string? alignment = null)
    {
        var paragraph = new Paragraph();
        var properties = new ParagraphProperties();
        if (styleId is not null)
            properties.Append(new ParagraphStyleId { Val = styleId });
        if (outlineLevel is int outline)
            properties.Append(new OutlineLevel { Val = outline });
        if (alignment == "center")
            properties.Append(new Justification { Val = JustificationValues.Center });
        if (properties.HasChildren)
            paragraph.Append(properties);
        paragraph.Append(first);
        body.Append(paragraph);
    }

    private void AddParagraph(OpenXmlElement[] children, string? styleId, int? outlineLevel)
    {
        var paragraph = new Paragraph();
        var properties = new ParagraphProperties();
        if (styleId is not null)
            properties.Append(new ParagraphStyleId { Val = styleId });
        if (outlineLevel is int outline)
            properties.Append(new OutlineLevel { Val = outline });
        if (properties.HasChildren)
            paragraph.Append(properties);
        foreach (var child in children)
            paragraph.Append(child);
        body.Append(paragraph);
    }

    private static Run WordRun(string text) =>
        new(new RunProperties(new Bold()), new Text(text) { Space = SpaceProcessingModeValues.Preserve });

    private void EnsureStyles()
    {
        var part = main.AddNewPart<StyleDefinitionsPart>();
        styles = new Styles();
        styles.InnerXml = """
            <w:docDefaults xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault>
              <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>
            </w:docDefaults>
            """;
        var normal = new Style { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true };
        normal.Append(new StyleName { Val = "Normal" });
        styles.Append(normal);
        part.Styles = styles;
    }

    private static AbstractNum Abstract(int identifier, NumberFormatValues format, string levelZeroText, string? levelOneText = null, NumberFormatValues? levelOneFormat = null, int start = 1)
    {
        var abstractNum = new AbstractNum { AbstractNumberId = identifier };
        abstractNum.Append(Level(0, format, levelZeroText, start));
        if (levelOneText is not null)
            abstractNum.Append(Level(1, levelOneFormat ?? format, levelOneText, 1));
        return abstractNum;
    }

    private static Level Level(int index, NumberFormatValues format, string text, int start) =>
        new(
            new StartNumberingValue { Val = start },
            new NumberingFormat { Val = format },
            new LevelText { Val = text })
        { LevelIndex = index };

    private static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");
}
