using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocxToHtml.Internal;
using DocxToHtml.Model;

namespace DocxToHtml.Reading;

internal sealed class ImageReader
{
    private static readonly Regex LengthPattern = new(
        @"(width|height)\s*:\s*([0-9.]+)\s*(pt|px|in|cm|mm)",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private readonly ReadingContext context;
    private readonly BlockReader blockReader;

    public ImageReader(ReadingContext context, BlockReader blockReader)
    {
        this.context = context;
        this.blockReader = blockReader;
    }

    public void ReadContainer(OpenXmlElement container, List<Inline> destination)
    {
        Walk(container, destination);
    }

    private void Walk(OpenXmlElement element, List<Inline> destination)
    {
        foreach (var child in MarkupCompatibility.Children(element))
        {
            if (child.LocalName == "txbxContent")
            {
                var blocks = blockReader.ReadBlocks(child);
                if (blocks.Count > 0)
                    destination.Add(new TextBoxInline { Blocks = blocks });
                continue;
            }

            if (child.LocalName == "blip")
            {
                AddBlip(child, destination);
                continue;
            }

            if (child.LocalName == "imagedata")
            {
                AddImageData(child, destination);
                continue;
            }

            Walk(child, destination);
        }
    }

    private void AddBlip(OpenXmlElement blip, List<Inline> destination)
    {
        var relationshipId = OpenXmlValues.Attribute(blip, "embed") ?? OpenXmlValues.Attribute(blip, "link");
        if (string.IsNullOrEmpty(relationshipId))
        {
            context.Warnings.Add("MissingImage", "A picture did not contain an image relationship.", context.Location);
            return;
        }

        var (width, height) = ReadExtent(blip);
        var image = context.Images.FindOrCreate(
            context.CurrentPart,
            relationshipId,
            ReadDescription(blip),
            width,
            height,
            context.Location);
        if (image is not null)
            destination.Add(image);
    }

    private void AddImageData(OpenXmlElement imageData, List<Inline> destination)
    {
        var relationshipId = OpenXmlValues.Attribute(imageData, "id");
        if (string.IsNullOrEmpty(relationshipId))
        {
            context.Warnings.Add("MissingImage", "A legacy picture did not contain an image relationship.", context.Location);
            return;
        }

        var (width, height) = ReadVmlSize(imageData);
        var title = OpenXmlValues.Attribute(imageData, "title") ?? "";
        var image = context.Images.FindOrCreate(
            context.CurrentPart,
            relationshipId,
            title,
            width,
            height,
            context.Location);
        if (image is not null)
            destination.Add(image);
    }

    private static (int? Width, int? Height) ReadExtent(OpenXmlElement element)
    {
        // wp:extent sits beside a:graphic, not above the blip, so it is not an ancestor.
        var drawing = NearestDrawing(element);
        if (drawing is null)
            return (null, null);

        foreach (var descendant in drawing.Descendants())
        {
            if (descendant.LocalName != "extent")
                continue;

            var horizontal = OpenXmlValues.LongAttribute(descendant, "cx");
            var vertical = OpenXmlValues.LongAttribute(descendant, "cy");
            return (
                horizontal is > 0 ? UnitConverter.EmuToPixels(horizontal.Value) : null,
                vertical is > 0 ? UnitConverter.EmuToPixels(vertical.Value) : null);
        }

        return (null, null);
    }

    private static string ReadDescription(OpenXmlElement element)
    {
        var drawing = NearestDrawing(element);
        if (drawing is null)
            return "";

        foreach (var descendant in drawing.Descendants())
        {
            if (descendant.LocalName is not ("docPr" or "cNvPr"))
                continue;

            var description = OpenXmlValues.Attribute(descendant, "descr");
            if (!string.IsNullOrWhiteSpace(description))
                return description;

            var title = OpenXmlValues.Attribute(descendant, "title");
            if (!string.IsNullOrWhiteSpace(title))
                return title;
        }

        return "";
    }

    private static OpenXmlElement? NearestDrawing(OpenXmlElement element)
    {
        var current = element.Parent;
        OpenXmlElement? drawing = null;
        while (current is not null)
        {
            if (current.LocalName is "inline" or "anchor" or "drawing" or "pict")
                drawing = current;

            current = current.Parent;
        }

        return drawing;
    }

    private static (int? Width, int? Height) ReadVmlSize(OpenXmlElement imageData)
    {
        var current = imageData.Parent;
        while (current is not null)
        {
            if (current.LocalName == "shape")
            {
                var style = OpenXmlValues.Attribute(current, "style");
                return (ParseLength(style, "width"), ParseLength(style, "height"));
            }

            current = current.Parent;
        }

        return (null, null);
    }

    private static int? ParseLength(string? style, string axis)
    {
        if (string.IsNullOrEmpty(style))
            return null;

        foreach (Match match in LengthPattern.Matches(style))
        {
            if (!match.Groups[1].Value.Equals(axis, StringComparison.OrdinalIgnoreCase))
                continue;

            if (!double.TryParse(match.Groups[2].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var magnitude))
                return null;

            var pixels = match.Groups[3].Value.ToLowerInvariant() switch
            {
                "px" => magnitude,
                "in" => magnitude * UnitConverter.CssPixelsPerInch,
                "cm" => magnitude / 2.54 * UnitConverter.CssPixelsPerInch,
                "mm" => magnitude / 25.4 * UnitConverter.CssPixelsPerInch,
                _ => magnitude * UnitConverter.CssPixelsPerInch / 72d
            };
            return (int)Math.Round(pixels, MidpointRounding.AwayFromZero);
        }

        return null;
    }
}
