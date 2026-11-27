using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocxToHtml.Internal;
using DocxToHtml.Model;

namespace DocxToHtml.Reading;

internal sealed class ImageCollector
{
    private readonly DocxToHtmlOptions options;
    private readonly WarningSink warnings;
    private readonly Dictionary<string, ImageInline> seen = new(StringComparer.Ordinal);
    private int nextNumber = 1;

    public ImageCollector(DocxToHtmlOptions options, WarningSink warnings)
    {
        this.options = options;
        this.warnings = warnings;
    }

    public List<ExtractedImage> Images { get; } = new();

    public ImageInline? FindOrCreate(
        OpenXmlPart sourcePart,
        string relationshipId,
        string altText,
        int? widthPixels,
        int? heightPixels,
        string? location)
    {
        if (options.ImageMode == ImageMode.Skip)
        {
            warnings.Add("ImageSkipped", "An image was skipped because ImageMode is Skip.", location);
            return null;
        }

        var cacheKey = sourcePart.Uri + "|" + relationshipId;
        if (seen.TryGetValue(cacheKey, out var existing))
        {
            return new ImageInline
            {
                FileName = existing.FileName,
                ContentType = existing.ContentType,
                Data = existing.Data,
                AltText = string.IsNullOrEmpty(altText) ? existing.AltText : altText,
                WidthPixels = widthPixels ?? existing.WidthPixels,
                HeightPixels = heightPixels ?? existing.HeightPixels
            };
        }

        IdPartPair? matched = null;
        foreach (var pair in sourcePart.Parts)
        {
            if (pair.RelationshipId == relationshipId)
            {
                matched = pair;
                break;
            }
        }

        if (matched?.OpenXmlPart is not ImagePart imagePart)
        {
            var external = sourcePart.ExternalRelationships.FirstOrDefault(relationship => relationship.Id == relationshipId);
            if (external is not null && options.ImageMode == ImageMode.External)
            {
                warnings.Add("MissingImage", "An image is linked externally and was not embedded in the package.", location);
                return new ImageInline
                {
                    FileName = external.Uri.IsAbsoluteUri ? external.Uri.AbsoluteUri : external.Uri.OriginalString,
                    ContentType = "application/octet-stream",
                    Data = Array.Empty<byte>(),
                    AltText = altText,
                    WidthPixels = widthPixels,
                    HeightPixels = heightPixels,
                    SourceIsUri = true
                };
            }

            warnings.Add("MissingImage", $"Image relationship '{relationshipId}' could not be read.", location);
            return null;
        }

        byte[] data;
        try
        {
            using var stream = imagePart.GetStream();
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            data = buffer.ToArray();
        }
        catch (IOException exception)
        {
            warnings.Add("MissingImage", "An image part could not be read: " + exception.Message, location);
            return null;
        }

        var contentType = string.IsNullOrWhiteSpace(imagePart.ContentType)
            ? "application/octet-stream"
            : imagePart.ContentType;
        var fileName = "image-" + nextNumber.ToString("000", CultureInfo.InvariantCulture) + "." + ExtensionFor(contentType);
        nextNumber++;

        var inline = new ImageInline
        {
            FileName = fileName,
            ContentType = contentType,
            Data = data,
            AltText = altText,
            WidthPixels = widthPixels,
            HeightPixels = heightPixels
        };
        seen[cacheKey] = inline;
        Images.Add(new ExtractedImage(fileName, contentType, data));
        return inline;
    }

    private static string ExtensionFor(string contentType) => contentType.ToLowerInvariant() switch
    {
        "image/png" => "png",
        "image/jpeg" or "image/jpg" => "jpg",
        "image/gif" => "gif",
        "image/bmp" or "image/x-bmp" => "bmp",
        "image/tiff" => "tiff",
        "image/webp" => "webp",
        "image/svg+xml" => "svg",
        "image/x-emf" or "image/emf" => "emf",
        "image/x-wmf" or "image/wmf" => "wmf",
        "image/x-icon" or "image/vnd.microsoft.icon" => "ico",
        _ => "bin"
    };
}
