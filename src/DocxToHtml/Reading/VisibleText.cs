using DocumentFormat.OpenXml;

namespace DocxToHtml.Reading;

internal static class VisibleText
{
    public static string From(OpenXmlElement element, TrackedChangesMode mode)
    {
        var builder = new StringBuilder();
        Walk(element, builder, mode);
        return builder.ToString();
    }

    private static void Walk(OpenXmlElement element, StringBuilder builder, TrackedChangesMode mode)
    {
        foreach (var child in Internal.MarkupCompatibility.Children(element))
        {
            if (child.LocalName is "ins" or "del" or "moveFrom" or "moveTo"
                && !RevisionFilter.Include(child.LocalName, mode))
            {
                continue;
            }

            if (child.LocalName is "instrText" or "delInstrText")
                continue;

            if (child.LocalName == "t")
                builder.Append(child.InnerText);
            else if (child.LocalName == "delText" && mode == TrackedChangesMode.Reject)
                builder.Append(child.InnerText);
            else
                Walk(child, builder, mode);
        }
    }
}
