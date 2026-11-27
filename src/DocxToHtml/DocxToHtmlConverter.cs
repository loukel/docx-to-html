using System.IO.Packaging;
using DocumentFormat.OpenXml.Packaging;
using DocxToHtml.Reading;
using DocxToHtml.Writing;

namespace DocxToHtml;

/// <summary>
/// Converts a Word <c>.docx</c> document into HTML and CSS.
/// The library does not write to the console. Only <see cref="ConvertToFolder"/> writes files.
/// </summary>
public static class DocxToHtmlConverter
{
    /// <summary>Converts the document at <paramref name="docxPath"/>.</summary>
    public static ConversionResult Convert(string docxPath, DocxToHtmlOptions? options = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(docxPath);
        if (!File.Exists(docxPath))
            throw new FileNotFoundException("The document was not found.", docxPath);

        try
        {
            using var wordDocument = WordprocessingDocument.Open(docxPath, false);
            return ConvertDocument(wordDocument, options ?? new DocxToHtmlOptions(), Path.GetFileNameWithoutExtension(docxPath));
        }
        catch (Exception exception) when (IsPackageException(exception))
        {
            throw new InvalidDataException("The file could not be read as a .docx document.", exception);
        }
    }

    /// <summary>
    /// Converts a readable document stream. The stream is copied, so the caller's stream is not disposed.
    /// The copy starts at the stream's current position.
    /// </summary>
    public static ConversionResult Convert(Stream docxStream, DocxToHtmlOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(docxStream);
        using var buffer = new MemoryStream();
        docxStream.CopyTo(buffer);
        buffer.Position = 0;
        try
        {
            using var wordDocument = WordprocessingDocument.Open(buffer, false);
            return ConvertDocument(wordDocument, options ?? new DocxToHtmlOptions(), sourceName: null);
        }
        catch (Exception exception) when (IsPackageException(exception))
        {
            throw new InvalidDataException("The file could not be read as a .docx document.", exception);
        }
    }

    /// <summary>
    /// Converts <paramref name="docxPath"/> and writes <c>index.html</c>, the stylesheet, and any images into <paramref name="outputDirectory"/>.
    /// </summary>
    public static ConversionResult ConvertToFolder(string docxPath, string outputDirectory, DocxToHtmlOptions? options = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(docxPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputDirectory);
        var effectiveOptions = options ?? new DocxToHtmlOptions();
        var result = Convert(docxPath, effectiveOptions);
        Directory.CreateDirectory(outputDirectory);
        File.WriteAllText(Path.Combine(outputDirectory, "index.html"), result.Html);
        File.WriteAllText(Path.Combine(outputDirectory, HtmlDocumentWriter.CssFileName(effectiveOptions)), result.Css);

        if (result.Images.Count > 0)
        {
            var imageDirectory = Path.Combine(outputDirectory, HtmlDocumentWriter.ImageFolder(effectiveOptions));
            Directory.CreateDirectory(imageDirectory);
            foreach (var image in result.Images)
            {
                var fileName = Path.GetFileName(image.FileName);
                if (string.IsNullOrEmpty(fileName))
                    continue;

                File.WriteAllBytes(Path.Combine(imageDirectory, fileName), image.Data);
            }
        }

        return result;
    }

    private static ConversionResult ConvertDocument(WordprocessingDocument wordDocument, DocxToHtmlOptions options, string? sourceName)
    {
        var model = new DocumentReader().Read(wordDocument, options, sourceName);
        return new HtmlDocumentWriter().Write(model, options);
    }

    private static bool IsPackageException(Exception exception) =>
        exception is FileFormatException
        || exception.GetType().Namespace?.StartsWith("DocumentFormat.OpenXml", StringComparison.Ordinal) == true;
}
