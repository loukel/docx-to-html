using System.Text;
using System.Text.RegularExpressions;
using AngleSharp.Dom;
using AngleSharp.Html.Parser;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxToHtml.Tests.Support;

namespace DocxToHtml.Tests;

public class ConversionTests
{
    public static IEnumerable<object[]> AllFixtures() =>
        WordFixtureFactory.Names.Select(name => new object[] { name });

    [Theory]
    [MemberData(nameof(AllFixtures))]
    public void EveryFixtureKeepsTextAndAvoidsInlineStyles(string fixtureName)
    {
        var documentBytes = WordFixtureFactory.Build(fixtureName);
        var options = new DocxToHtmlOptions();
        var result = ConvertBytes(documentBytes, options);
        var document = new HtmlParser().ParseDocument(result.Html);

        Assert.DoesNotContain("style=", result.Html, StringComparison.OrdinalIgnoreCase);
        foreach (var element in document.All)
            Assert.True(element.GetAttribute("style") is null, element.OuterHtml);

        foreach (var span in document.QuerySelectorAll("span"))
            Assert.False(string.IsNullOrWhiteSpace(span.ClassName), span.OuterHtml);

        foreach (var heading in document.QuerySelectorAll("h1, h2, h3, h4, h5, h6"))
            Assert.True(int.TryParse(heading.LocalName[1..], out var level) && level is >= 1 and <= 6);

        var expected = VisibleDocumentText.Extract(documentBytes, options);
        foreach (var generated in document.QuerySelectorAll("[class*='footnote-ref'], [class*='footnote-backlink']"))
            generated.Remove();
        Assert.Equal(expected, VisibleHtmlText(document.Body));
    }

    [Fact]
    public void CollectsEveryHeadingAndMarksLevelsAboveSix()
    {
        var result = ConvertBytes(WordFixtureFactory.Build("headings"));
        var document = new HtmlParser().ParseDocument(result.Html);

        Assert.Equal(
            [1, 2, 3, 4, 5, 6, 7, 8, 9, 1, 2, 7, 1],
            result.Headings.Select(heading => heading.Level).ToArray());
        Assert.Equal(
            ["Chapter", "Section", "Topic", "Detail", "Note", "Aside", "Deep", "Deeper", "Deepest", "Document title", "Appendix", "Buried", "Outlined"],
            result.Headings.Select(heading => heading.Text).ToArray());
        Assert.DoesNotContain(result.Headings, heading => heading.Text == "A subtitle");
        Assert.DoesNotContain(result.Headings, heading => heading.Text == "Body text");

        Assert.Null(document.QuerySelector("h7, h8, h9"));
        var deep = document.QuerySelector("h6.p-heading-7.h7");
        Assert.NotNull(deep);
        Assert.Equal("Deep", deep!.TextContent);
        Assert.NotNull(document.QuerySelector("h6.h8"));
        Assert.NotNull(document.QuerySelector("h6.h9"));
        Assert.NotNull(document.QuerySelector("h1.title"));
        Assert.NotNull(document.QuerySelector("p.subtitle"));
        Assert.NotNull(document.QuerySelector("h2.p-appendix-heading"));
        Assert.NotNull(document.QuerySelector("h6.p-deep-custom.h7"));
        Assert.DoesNotContain("<strong>Chapter</strong>", result.Html, StringComparison.Ordinal);
    }

    [Fact]
    public void PrefixesGeneratedHeadingClasses()
    {
        var result = ConvertBytes(
            WordFixtureFactory.Build("headings"),
            new DocxToHtmlOptions { ClassPrefix = "dx-" });
        var document = new HtmlParser().ParseDocument(result.Html);
        Assert.NotNull(document.QuerySelector("h6.dx-h7"));
        Assert.NotNull(document.QuerySelector("h6.dx-p-heading-7"));
        Assert.Contains(".dx-p-heading-1", result.Css, StringComparison.Ordinal);
    }

    [Fact]
    public void StyleMapOverridesHeadingDetection()
    {
        var options = new DocxToHtmlOptions
        {
            StyleMap = { ["heading 1"] = new StyleMapping("p", "lead") }
        };
        var result = ConvertBytes(WordFixtureFactory.Build("headings"), options);
        Assert.DoesNotContain(result.Headings, heading => heading.Text == "Chapter");
        Assert.Contains("class=\"lead\"", result.Html, StringComparison.Ordinal);
        Assert.Contains(result.Headings, heading => heading.Text == "Section" && heading.Level == 2);
    }

    [Fact]
    public void DirectFormattingModesChangeRunMarkup()
    {
        var bytes = WordFixtureFactory.Build("run-formatting");
        var none = ConvertBytes(bytes, new DocxToHtmlOptions { DirectFormatting = DirectFormattingMode.None });
        var semantic = ConvertBytes(bytes, new DocxToHtmlOptions { DirectFormatting = DirectFormattingMode.SemanticOnly });
        var full = ConvertBytes(bytes, new DocxToHtmlOptions { DirectFormatting = DirectFormattingMode.SemanticAndClasses });

        Assert.DoesNotContain("<strong>", none.Html, StringComparison.Ordinal);
        Assert.Contains("<strong>Bold words</strong>", semantic.Html, StringComparison.Ordinal);
        Assert.DoesNotContain("color-", semantic.Html, StringComparison.Ordinal);
        Assert.Contains("<strong>Bold words</strong>", full.Html, StringComparison.Ordinal);
        Assert.Contains("color-ff0000", full.Html, StringComparison.Ordinal);
        Assert.Contains("highlight-ffff00", full.Html, StringComparison.Ordinal);
        Assert.DoesNotContain("style=", full.Html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MergesFragmentedRuns()
    {
        var result = ConvertBytes(WordFixtureFactory.Build("fragmented-runs"));
        Assert.Contains("<strong>Hello</strong>", result.Html, StringComparison.Ordinal);
        Assert.Single(Regex.Matches(result.Html, "<strong>"));
    }

    [Fact]
    public void HtmlWrittenToTheFileIsIndented()
    {
        var result = ConvertBytes(WordFixtureFactory.Build("lists"));
        Assert.StartsWith("<!DOCTYPE html>\n<html lang=\"en-US\">\n  <head>\n    <meta charset=\"utf-8\">\n", result.Html, StringComparison.Ordinal);
        Assert.Contains("\n  <body>\n    <ul>\n      <li>Fruit<ul>\n          <li>Apple</li>\n", result.Html, StringComparison.Ordinal);
        Assert.Contains("\n    <ol start=\"5\">\n      <li>Fifth</li>\n    </ol>\n", result.Html, StringComparison.Ordinal);
        Assert.EndsWith("</body>\n</html>\n", result.Html, StringComparison.Ordinal);
    }

    [Fact]
    public void NumberedHeadingsKeepTheNumberInsideTheHeading()
    {
        var builder = new WordDocumentBuilder("Numbered headings");
        builder.AddMultiLevelHeadingNumbering();
        builder.AddParagraph("Scope", "Heading1");
        builder.AddParagraph("General", "Heading2");
        builder.AddParagraph("More", "Heading3");
        builder.AddParagraph("Deep", "Heading4");
        builder.AddParagraph("Other", "Heading2");
        builder.AddParagraph("", "Heading1");
        var result = ConvertBytes(builder.Finish());
        var document = new HtmlParser().ParseDocument(result.Html);

        Assert.Empty(document.QuerySelectorAll("ol, ul"));
        Assert.Equal(
            ["1 Scope", "1.1 General", "1.1.1 More", "1.1.1.1 Deep", "1.2 Other"],
            result.Headings.Select(heading => heading.Text).ToArray());
        Assert.DoesNotContain(document.QuerySelectorAll("h1, h2, h3, h4, h5, h6"), heading => string.IsNullOrWhiteSpace(heading.TextContent));
        Assert.Contains(">1.1.1.1 Deep</h4>", result.Html, StringComparison.Ordinal);
    }

    [Fact]
    public void ContentsLinksFindHeadingBookmarks()
    {
        var builder = new WordDocumentBuilder("Contents");
        builder.AddParagraphStyle("Heading1", "heading 1", outlineLevel: 0, bold: true);
        builder.Add(new Paragraph(
            new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
            new BookmarkStart { Id = "7", Name = "_Toc12345" },
            new Run(new Text("Scope")),
            new BookmarkEnd { Id = "7" }));
        builder.Add(new Paragraph(
            new Hyperlink(new Run(new Text("Scope"))) { Anchor = "_Toc12345" }));
        var result = ConvertBytes(builder.Finish());
        var document = new HtmlParser().ParseDocument(result.Html);

        var heading = document.QuerySelector("h1");
        Assert.NotNull(heading);
        Assert.Equal("_Toc12345", heading!.Id);
        Assert.Equal("#_Toc12345", document.QuerySelector("a[href]")?.GetAttribute("href"));
    }

    [Fact]
    public void HyperlinksAndBookmarksBecomeAnchors()
    {
        var result = ConvertBytes(WordFixtureFactory.Build("hyperlinks"));
        Assert.Contains("href=\"https://example.com/\"", result.Html, StringComparison.Ordinal);
        Assert.Contains("id=\"chapter\"", result.Html, StringComparison.Ordinal);
        Assert.Contains("href=\"#chapter\"", result.Html, StringComparison.Ordinal);
    }

    [Fact]
    public void FieldsKeepResultsAndDropInstructions()
    {
        var result = ConvertBytes(WordFixtureFactory.Build("fields"));
        Assert.Contains("href=\"https://example.com/complex\"", result.Html, StringComparison.Ordinal);
        Assert.Contains("href=\"https://example.com/simple\"", result.Html, StringComparison.Ordinal);
        Assert.Contains("27 November 2026", result.Html, StringComparison.Ordinal);
        Assert.DoesNotContain("DATE", result.Html, StringComparison.Ordinal);
        Assert.DoesNotContain("HYPERLINK", result.Html, StringComparison.Ordinal);
    }

    [Fact]
    public void TrackedChangesFollowTheSelectedMode()
    {
        var bytes = WordFixtureFactory.Build("tracked-changes");
        var accepted = ConvertBytes(bytes, new DocxToHtmlOptions { TrackedChanges = TrackedChangesMode.Accept });
        var rejected = ConvertBytes(bytes, new DocxToHtmlOptions { TrackedChanges = TrackedChangesMode.Reject });
        Assert.Contains("added", accepted.Html, StringComparison.Ordinal);
        Assert.DoesNotContain("removed", accepted.Html, StringComparison.Ordinal);
        Assert.Contains("removed", rejected.Html, StringComparison.Ordinal);
        Assert.DoesNotContain("added", rejected.Html, StringComparison.Ordinal);
    }

    [Fact]
    public void ListsUseTheNumberingDefinition()
    {
        var document = new HtmlParser().ParseDocument(ConvertBytes(WordFixtureFactory.Build("lists")).Html);
        Assert.Equal(2, document.QuerySelectorAll("ul").Length);
        Assert.Equal("5", document.QuerySelector("ol[start]")?.GetAttribute("start"));
        Assert.NotNull(document.QuerySelector("ol.lower-alpha"));
    }

    [Fact]
    public void TablesExposeSpansHeaderShadingAndNesting()
    {
        var spanned = new HtmlParser().ParseDocument(ConvertBytes(WordFixtureFactory.Build("colspan-rowspan")).Html);
        var merged = spanned.QuerySelector("td[colspan='2'][rowspan='2']");
        Assert.NotNull(merged);
        Assert.Equal("Ada", merged!.TextContent);
        Assert.Equal(2, spanned.QuerySelectorAll("tr").Length);

        var header = new HtmlParser().ParseDocument(ConvertBytes(WordFixtureFactory.Build("header-row")).Html);
        Assert.NotNull(header.QuerySelector("thead th"));
        Assert.Equal("Ada", header.QuerySelector("tbody td")?.TextContent);

        var simple = new HtmlParser().ParseDocument(ConvertBytes(WordFixtureFactory.Build("simple-table")).Html);
        Assert.NotNull(simple.QuerySelector("td.bg-d9e2f3"));
        Assert.NotNull(simple.QuerySelector("td.align-center"));

        var nested = new HtmlParser().ParseDocument(ConvertBytes(WordFixtureFactory.Build("nested-table")).Html);
        Assert.Equal(2, nested.QuerySelectorAll("table").Length);
        Assert.Contains("Inner", nested.QuerySelector("table table")?.TextContent, StringComparison.Ordinal);

        var wrapped = new HtmlParser().ParseDocument(ConvertBytes(WordFixtureFactory.Build("table-in-content-control")).Html);
        Assert.Contains("Inside", wrapped.QuerySelector("td")?.TextContent, StringComparison.Ordinal);
    }

    [Fact]
    public void ImagesUseClassesAndAltText()
    {
        var result = ConvertBytes(WordFixtureFactory.Build("images"));
        var image = new HtmlParser().ParseDocument(result.Html).QuerySelector("img");
        Assert.NotNull(image);
        Assert.Equal("Sample chart", image!.GetAttribute("alt"));
        Assert.Null(image.GetAttribute("width"));
        Assert.Null(image.GetAttribute("height"));
        Assert.Contains("img-w-320-h-160", image.ClassName, StringComparison.Ordinal);
        Assert.Contains("images/image-001.png", image.GetAttribute("src"), StringComparison.Ordinal);
        Assert.Single(result.Images);
        Assert.Equal("image-001.png", result.Images[0].FileName);

        var embedded = ConvertBytes(WordFixtureFactory.Build("images"), new DocxToHtmlOptions { ImageMode = ImageMode.DataUri });
        Assert.Contains("data:image/png;base64,", embedded.Html, StringComparison.Ordinal);

        var skipped = ConvertBytes(WordFixtureFactory.Build("images"), new DocxToHtmlOptions { ImageMode = ImageMode.Skip });
        Assert.DoesNotContain("<img", skipped.Html, StringComparison.Ordinal);
        Assert.Contains(skipped.Warnings, warning => warning.Code == "ImageSkipped");
    }

    [Fact]
    public void TextBoxesAreNotInsideParagraphs()
    {
        var document = new HtmlParser().ParseDocument(ConvertBytes(WordFixtureFactory.Build("text-box")).Html);
        var box = document.QuerySelector("div.textbox");
        Assert.NotNull(box);
        Assert.Equal("Inside the box", VisibleDocumentText.Normalize(box!.TextContent));
        Assert.NotEqual("P", box.ParentElement?.TagName, StringComparer.OrdinalIgnoreCase);
        Assert.Contains("Label", document.Body!.TextContent, StringComparison.Ordinal);
    }

    [Fact]
    public void FootnotesStayOutUntilRequestedAndHeadersStayOutByDefault()
    {
        var bytes = WordFixtureFactory.Build("footnotes");
        var shown = ConvertBytes(bytes);
        Assert.Contains("A note about sources", shown.Html, StringComparison.Ordinal);
        Assert.Contains("class=\"footnote-ref\"", shown.Html, StringComparison.Ordinal);
        Assert.Contains("class=\"footnote-backlink\"", shown.Html, StringComparison.Ordinal);
        Assert.DoesNotContain("Running header", shown.Html, StringComparison.Ordinal);

        var hidden = ConvertBytes(bytes, new DocxToHtmlOptions { IncludeFootnotes = false });
        Assert.DoesNotContain("A note about sources", hidden.Html, StringComparison.Ordinal);

        var withChrome = ConvertBytes(bytes, new DocxToHtmlOptions { IncludeHeadersAndFooters = true });
        Assert.Contains("Running header", withChrome.Html, StringComparison.Ordinal);
        Assert.Contains("Running footer", withChrome.Html, StringComparison.Ordinal);
    }

    [Fact]
    public void OutputModesAndFolderWriter()
    {
        var bytes = WordFixtureFactory.Build("paragraph-styles");
        var fragment = ConvertBytes(bytes, new DocxToHtmlOptions { OutputMode = HtmlOutputMode.BodyFragment });
        Assert.DoesNotContain("<!DOCTYPE", fragment.Html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("A quoted line.", fragment.Html, StringComparison.Ordinal);

        var embedded = ConvertBytes(bytes, new DocxToHtmlOptions { CssMode = CssMode.Embedded });
        Assert.Contains("<style>", embedded.Html, StringComparison.Ordinal);
        Assert.Contains("A quoted line.", embedded.Html, StringComparison.Ordinal);

        var directory = Directory.CreateTempSubdirectory("docxtohtml");
        var documentPath = Path.Combine(directory.FullName, "source.docx");
        try
        {
            File.WriteAllBytes(documentPath, bytes);
            var written = DocxToHtmlConverter.ConvertToFolder(documentPath, directory.FullName);
            Assert.True(File.Exists(Path.Combine(directory.FullName, "index.html")));
            Assert.True(File.Exists(Path.Combine(directory.FullName, "styles.css")));
            Assert.Equal(written.Css, File.ReadAllText(Path.Combine(directory.FullName, "styles.css")));
            Assert.Contains("Just a sentence.", File.ReadAllText(Path.Combine(directory.FullName, "index.html")), StringComparison.Ordinal);
        }
        finally
        {
            directory.Delete(true);
        }
    }

    [Fact]
    public void EmptyParagraphsAreOmittedUnlessRequested()
    {
        var result = ConvertBytes(WordFixtureFactory.Build("paragraph-styles"), new DocxToHtmlOptions { IncludeEmptyParagraphs = false });
        Assert.DoesNotContain("<p></p>", result.Html, StringComparison.Ordinal);
        Assert.DoesNotContain("<p>\n</p>", result.Html, StringComparison.Ordinal);

        var kept = ConvertBytes(WordFixtureFactory.Build("paragraph-styles"), new DocxToHtmlOptions { IncludeEmptyParagraphs = true });
        Assert.Contains("<p></p>", kept.Html, StringComparison.Ordinal);
    }

    [Fact]
    public void InvalidPackageThrows()
    {
        using var stream = new MemoryStream("not a docx"u8.ToArray());
        Assert.Throws<InvalidDataException>(() => DocxToHtmlConverter.Convert(stream));
    }

    [Theory]
    [MemberData(nameof(AllFixtures))]
    public void SnapshotsMatch(string fixtureName)
    {
        var result = ConvertBytes(WordFixtureFactory.Build(fixtureName));
        var directory = Path.Combine(TestPaths.ProjectDirectory(), "Snapshots");
        Directory.CreateDirectory(directory);
        var htmlPath = Path.Combine(directory, fixtureName + ".html");
        var cssPath = Path.Combine(directory, fixtureName + ".css");
        if (Environment.GetEnvironmentVariable("UPDATE_SNAPSHOTS") == "1")
        {
            File.WriteAllText(htmlPath, result.Html);
            File.WriteAllText(cssPath, result.Css);
            var fixtureDirectory = Path.Combine(TestPaths.ProjectDirectory(), "Fixtures");
            Directory.CreateDirectory(fixtureDirectory);
            File.WriteAllBytes(Path.Combine(fixtureDirectory, fixtureName + ".docx"), WordFixtureFactory.Build(fixtureName));
            return;
        }

        Assert.True(File.Exists(htmlPath), $"Missing snapshot {htmlPath}. Set UPDATE_SNAPSHOTS=1 to create it.");
        Assert.Equal(File.ReadAllText(htmlPath), result.Html);
        Assert.Equal(File.ReadAllText(cssPath), result.Css);
    }

    private static string VisibleHtmlText(IElement? body)
    {
        if (body is null)
            return "";

        var builder = new StringBuilder();
        AppendVisibleText(body, builder);
        return VisibleDocumentText.Normalize(builder.ToString());
    }

    private static void AppendVisibleText(INode node, StringBuilder builder)
    {
        foreach (var child in node.ChildNodes)
        {
            if (child is IText textNode)
            {
                if (!string.IsNullOrWhiteSpace(textNode.Data.Replace('\u00A0', ' ')))
                    builder.Append(textNode.Data);
            }
            else
            {
                AppendVisibleText(child, builder);
            }
        }
    }

    private static ConversionResult ConvertBytes(byte[] documentBytes, DocxToHtmlOptions? options = null)
    {
        using var stream = new MemoryStream(documentBytes);
        return DocxToHtmlConverter.Convert(stream, options);
    }
}
