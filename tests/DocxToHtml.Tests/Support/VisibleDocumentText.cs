using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocxToHtml;

namespace DocxToHtml.Tests.Support;

/// <summary>
/// Independent text walk used by the completeness invariant.
/// It is intentionally not the converter's reader.
/// </summary>
internal static class VisibleDocumentText
{
    public static string Extract(byte[] documentBytes, DocxToHtmlOptions options)
    {
        using var stream = new MemoryStream(documentBytes);
        using var wordDocument = WordprocessingDocument.Open(stream, false);
        var main = wordDocument.MainDocumentPart ?? throw new InvalidDataException("Missing main part.");
        var builder = new StringBuilder();
        var footnoteIdentifiers = new List<int>();
        var endnoteIdentifiers = new List<int>();
        if (main.Document?.Body is not null)
            Walk(main.Document.Body, builder, options, footnoteIdentifiers, endnoteIdentifiers);

        if (options.IncludeHeadersAndFooters)
        {
            foreach (var headerPart in main.HeaderParts)
            {
                if (headerPart.Header is not null)
                    Walk(headerPart.Header, builder, options, footnoteIdentifiers, endnoteIdentifiers);
            }

            foreach (var footerPart in main.FooterParts)
            {
                if (footerPart.Footer is not null)
                    Walk(footerPart.Footer, builder, options, footnoteIdentifiers, endnoteIdentifiers);
            }
        }

        if (options.IncludeFootnotes)
        {
            AppendNotes(main.FootnotesPart?.Footnotes, footnoteIdentifiers, builder, options);
            AppendNotes(main.EndnotesPart?.Endnotes, endnoteIdentifiers, builder, options);
        }

        return Normalize(builder.ToString());
    }

    public static string Normalize(string? text)
    {
        if (string.IsNullOrEmpty(text))
            return "";

        var builder = new StringBuilder(text.Length);
        var pendingSpace = false;
        foreach (var character in text)
        {
            if (char.IsWhiteSpace(character) || character == '\u00A0' || character == '\u2011' || character == '\u00AD')
            {
                pendingSpace = builder.Length > 0;
                continue;
            }

            if (pendingSpace)
            {
                builder.Append(' ');
                pendingSpace = false;
            }

            builder.Append(character);
        }

        return builder.ToString();
    }

    private static void AppendNotes(OpenXmlElement? root, IReadOnlyList<int> identifiers, StringBuilder builder, DocxToHtmlOptions options)
    {
        if (root is null)
            return;

        foreach (var identifier in identifiers.Distinct())
        {
            foreach (var note in root.ChildElements)
            {
                if (note.LocalName is not ("footnote" or "endnote"))
                    continue;

                var noteIdentifier = Attribute(note, "id");
                if (noteIdentifier != identifier.ToString(System.Globalization.CultureInfo.InvariantCulture))
                    continue;

                var type = Attribute(note, "type");
                if (identifier <= 0 || type is "separator" or "continuationSeparator" or "continuationNotice")
                    continue;

                Walk(note, builder, options, new List<int>(), new List<int>());
            }
        }
    }

    private static void Walk(
        OpenXmlElement element,
        StringBuilder builder,
        DocxToHtmlOptions options,
        List<int> footnoteIdentifiers,
        List<int> endnoteIdentifiers)
    {
        foreach (var child in Children(element))
        {
            if (!IncludeRevision(child.LocalName, options.TrackedChanges))
                continue;

            if (child.LocalName is "instrText" or "delInstrText")
                continue;

            if (child.LocalName == "t")
            {
                builder.Append(child.InnerText);
                continue;
            }

            if (child.LocalName == "delText" && options.TrackedChanges == TrackedChangesMode.Reject)
            {
                builder.Append(child.InnerText);
                continue;
            }

            if (child.LocalName == "footnoteReference" && options.IncludeFootnotes && int.TryParse(Attribute(child, "id"), out var footnoteIdentifier))
                footnoteIdentifiers.Add(footnoteIdentifier);
            else if (child.LocalName == "endnoteReference" && options.IncludeFootnotes && int.TryParse(Attribute(child, "id"), out var endnoteIdentifier))
                endnoteIdentifiers.Add(endnoteIdentifier);

            if (child.LocalName is "mc:AlternateContent")
                continue;

            Walk(child, builder, options, footnoteIdentifiers, endnoteIdentifiers);
        }
    }

    private static IEnumerable<OpenXmlElement> Children(OpenXmlElement parent)
    {
        foreach (var child in parent.ChildElements)
        {
            if (child is AlternateContent alternateContent)
            {
                var branch = (OpenXmlElement?)alternateContent.GetFirstChild<AlternateContentChoice>()
                    ?? alternateContent.GetFirstChild<AlternateContentFallback>();
                if (branch is null)
                    continue;

                foreach (var nested in Children(branch))
                    yield return nested;
            }
            else
            {
                yield return child;
            }
        }
    }

    private static bool IncludeRevision(string localName, TrackedChangesMode mode)
    {
        var accepting = mode != TrackedChangesMode.Reject;
        return localName switch
        {
            "ins" or "moveTo" => accepting,
            "del" or "moveFrom" => !accepting,
            _ => true
        };
    }

    private static string? Attribute(OpenXmlElement element, string localName)
    {
        foreach (var attribute in element.GetAttributes())
        {
            if (attribute.LocalName == localName)
                return attribute.Value;
        }

        return null;
    }
}
