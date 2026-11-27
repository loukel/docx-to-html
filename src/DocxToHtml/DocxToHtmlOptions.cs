namespace DocxToHtml;

/// <summary>How much of an HTML document to return.</summary>
public enum HtmlOutputMode
{
    /// <summary>A complete document with a doctype, head, and body.</summary>
    FullDocument,

    /// <summary>Only the contents of the body. No doctype and no head.</summary>
    BodyFragment
}

/// <summary>Where generated CSS is referenced from the HTML.</summary>
public enum CssMode
{
    /// <summary>The HTML links to <see cref="DocxToHtmlOptions.CssFileName"/>.</summary>
    External,

    /// <summary>The stylesheet is placed in a <c>style</c> element in the document head.</summary>
    Embedded,

    /// <summary>The HTML does not reference CSS. <see cref="ConversionResult.Css"/> is still populated.</summary>
    None
}

/// <summary>How run and paragraph direct formatting is written.</summary>
public enum DirectFormattingMode
{
    /// <summary>Drop direct formatting. Paragraph styles, headings, lists, and tables remain.</summary>
    None,

    /// <summary>Emit <c>strong</c>, <c>em</c>, <c>u</c>, <c>s</c>, <c>sub</c>, and <c>sup</c> only.</summary>
    SemanticOnly,

    /// <summary>Semantic tags plus classes for colour, highlight, size, font, and direct paragraph layout.</summary>
    SemanticAndClasses
}

/// <summary>How embedded pictures are referenced.</summary>
public enum ImageMode
{
    /// <summary>Write <c>src</c> as a path under <see cref="DocxToHtmlOptions.ImageFolder"/>.</summary>
    External,

    /// <summary>Write <c>src</c> as a base64 data URI. Bytes are still returned on the result.</summary>
    DataUri,

    /// <summary>Omit the image and record an <c>ImageSkipped</c> warning.</summary>
    Skip
}

/// <summary>Which side of tracked changes becomes the visible text.</summary>
public enum TrackedChangesMode
{
    /// <summary>Keep insertions and move destinations. Drop deletions and move sources.</summary>
    Accept,

    /// <summary>Keep deletions and move sources. Drop insertions and move destinations.</summary>
    Reject
}

/// <summary>
/// Forces one Word paragraph style onto a specific HTML tag.
/// Lookup is by style name, not style id.
/// </summary>
/// <param name="Tag">HTML tag name, for example <c>blockquote</c> or <c>h2</c>.</param>
/// <param name="CssClass">
/// Optional class token or space-separated tokens. These are not passed through
/// <see cref="DocxToHtmlOptions.ClassPrefix"/>.
/// </param>
public sealed record StyleMapping(string Tag, string? CssClass = null);

/// <summary>Settings for a single conversion. The converter does not modify the instance.</summary>
public sealed class DocxToHtmlOptions
{
    /// <summary>Full document or body fragment. Defaults to a full document.</summary>
    public HtmlOutputMode OutputMode { get; set; } = HtmlOutputMode.FullDocument;

    /// <summary>
    /// Title used in the document head. When null, the core-properties title is used,
    /// then the file name, then <c>Document</c>.
    /// </summary>
    public string? DocumentTitle { get; set; }

    /// <summary>How the HTML refers to the stylesheet. Defaults to an external link.</summary>
    public CssMode CssMode { get; set; } = CssMode.External;

    /// <summary>File name used by the external stylesheet link and by <see cref="DocxToHtmlConverter.ConvertToFolder"/>.</summary>
    public string CssFileName { get; set; } = "styles.css";

    /// <summary>
    /// Prefix added to generated class names. <c>dx-</c> turns <c>p-heading-1</c> into <c>dx-p-heading-1</c>
    /// and a level-7 heading class into <c>dx-h7</c>.
    /// </summary>
    public string ClassPrefix { get; set; } = "";

    /// <summary>CSS appended after the generated rules. Not parsed or rewritten.</summary>
    public string? AdditionalCss { get; set; }

    /// <summary>Include the built-in rules for body text, tables, lists, notes, and text boxes.</summary>
    public bool IncludeBaseCss { get; set; } = true;

    /// <summary>How direct formatting is turned into tags and classes.</summary>
    public DirectFormattingMode DirectFormatting { get; set; } = DirectFormattingMode.SemanticAndClasses;

    /// <summary>Map Heading styles and outline levels onto <c>h1</c>–<c>h6</c>.</summary>
    public bool UseSemanticHeadings { get; set; } = true;

    /// <summary>
    /// Overrides heading detection for the named Word styles.
    /// The key is the style name, compared without case sensitivity.
    /// </summary>
    public Dictionary<string, StyleMapping> StyleMap { get; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>When false, paragraphs with no visible content are omitted. List items are kept.</summary>
    public bool IncludeEmptyParagraphs { get; set; } = false;

    /// <summary>How pictures are referenced from the HTML.</summary>
    public ImageMode ImageMode { get; set; } = ImageMode.External;

    /// <summary>Relative folder used in external image paths and by <see cref="DocxToHtmlConverter.ConvertToFolder"/>.</summary>
    public string ImageFolder { get; set; } = "images";

    /// <summary>Include the default header and footer of the first section.</summary>
    public bool IncludeHeadersAndFooters { get; set; } = false;

    /// <summary>Include footnotes and endnotes.</summary>
    public bool IncludeFootnotes { get; set; } = true;

    /// <summary>Include comment text at each comment reference.</summary>
    public bool IncludeComments { get; set; } = false;

    /// <summary>Which side of tracked changes is visible.</summary>
    public TrackedChangesMode TrackedChanges { get; set; } = TrackedChangesMode.Accept;

    /// <summary>Called as each warning is recorded. Exceptions from the callback are not caught.</summary>
    public Action<ConversionWarning>? OnWarning { get; set; }
}
