namespace DocxToHtml;

/// <summary>
/// HTML, stylesheet, images, headings, and warnings produced from one document.
/// </summary>
public sealed class ConversionResult
{
    /// <summary>Full document or body fragment, depending on <see cref="DocxToHtmlOptions.OutputMode"/>.</summary>
    public required string Html { get; init; }

    /// <summary>Stylesheet text. Still returned when the HTML does not reference it.</summary>
    public required string Css { get; init; }

    /// <summary>Images extracted from the document, in first-seen order.</summary>
    public required IReadOnlyList<ExtractedImage> Images { get; init; }

    /// <summary>Recoverable problems encountered while converting. Nothing here dropped visible text on purpose.</summary>
    public required IReadOnlyList<ConversionWarning> Warnings { get; init; }

    /// <summary>
    /// Every heading in the order it appears in the HTML.
    /// <see cref="CollectedHeading.Level"/> is the Word level and is not clamped at 6.
    /// </summary>
    public required IReadOnlyList<CollectedHeading> Headings { get; init; }
}

/// <summary>An image part copied out of the package.</summary>
/// <param name="FileName">File name such as <c>image-001.png</c>. Does not include a folder.</param>
/// <param name="ContentType">MIME type from the image part.</param>
/// <param name="Data">Raw bytes of the image.</param>
public sealed record ExtractedImage(string FileName, string ContentType, byte[] Data);

/// <summary>A recoverable conversion problem.</summary>
/// <param name="Code">Stable identifier such as <c>UnknownElement</c> or <c>MissingImage</c>.</param>
/// <param name="Message">What happened.</param>
/// <param name="Location">Where in the document, when that is known.</param>
public sealed record ConversionWarning(string Code, string Message, string? Location);

/// <summary>
/// One heading from the document.
/// Levels above 6 are preserved here. In the HTML they are written as <c>h6</c>
/// with an additional class <c>h{Level}</c> (prefixed when <see cref="DocxToHtmlOptions.ClassPrefix"/> is set).
/// </summary>
/// <param name="Level">1-based heading level. Values above 6 are kept.</param>
/// <param name="Text">Visible text of the heading, with whitespace collapsed.</param>
/// <param name="AnchorId">HTML id of the heading element when a bookmark supplied one.</param>
public sealed record CollectedHeading(int Level, string Text, string? AnchorId);
