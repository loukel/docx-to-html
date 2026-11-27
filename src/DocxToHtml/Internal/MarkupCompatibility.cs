using DocumentFormat.OpenXml;

namespace DocxToHtml.Internal;

internal static class MarkupCompatibility
{
    /// <summary>
    /// Yields children in order. Alternate content contributes only the Choice branch,
    /// or the Fallback when there is no Choice, so drawings are not emitted twice.
    /// </summary>
    public static IEnumerable<OpenXmlElement> Children(OpenXmlElement parent)
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
}
