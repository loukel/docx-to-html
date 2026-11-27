using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxToHtml.Internal;

namespace DocxToHtml.Reading;

internal sealed class CommentLookup
{
    private readonly Dictionary<string, string> textByIdentifier = new(StringComparer.Ordinal);

    public static CommentLookup Load(WordprocessingDocument wordDocument)
    {
        var lookup = new CommentLookup();
        var comments = wordDocument.MainDocumentPart?.WordprocessingCommentsPart?.Comments;
        if (comments is null)
            return lookup;

        foreach (var comment in comments.Elements<Comment>())
        {
            var identifier = comment.Id?.Value;
            if (string.IsNullOrEmpty(identifier))
                continue;

            lookup.textByIdentifier[identifier] = HtmlText.CollapseWhitespace(comment.InnerText);
        }

        return lookup;
    }

    public string? Find(string identifier) =>
        textByIdentifier.TryGetValue(identifier, out var text) ? text : null;
}
