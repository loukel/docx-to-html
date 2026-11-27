using DocxToHtml.Internal;
using DocxToHtml.Model;

namespace DocxToHtml.Writing;

internal sealed class CssBuilder
{
    private readonly DocumentModel document;
    private readonly DocxToHtmlOptions options;
    private readonly ClassRegistry registry;
    private readonly IReadOnlyList<UsedStyle> usedStyles;

    public CssBuilder(DocumentModel document, DocxToHtmlOptions options, ClassRegistry registry, IReadOnlyList<UsedStyle> usedStyles)
    {
        this.document = document;
        this.options = options;
        this.registry = registry;
        this.usedStyles = usedStyles;
    }

    public string Build()
    {
        var builder = new StringBuilder();
        if (options.IncludeBaseCss)
            builder.Append(BaseRules());

        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var used in usedStyles)
        {
            var key = (used.IsTable ? "t:" : "p:") + used.ClassName;
            if (!seen.Add(key))
                continue;

            if (used.IsTable)
                AppendTableStyle(builder, used);
            else
                AppendParagraphStyle(builder, used);
        }

        foreach (var registered in registry.Classes)
        {
            builder.Append('.').Append(registered.ClassName).Append(" {\n");
            foreach (var declaration in registered.Declarations)
                builder.Append("  ").Append(declaration).Append('\n');
            builder.Append("}\n");
        }

        if (!string.IsNullOrWhiteSpace(options.AdditionalCss))
        {
            builder.Append(options.AdditionalCss.Trim());
            if (!options.AdditionalCss.TrimEnd().EndsWith('\n'))
                builder.Append('\n');
        }

        return builder.ToString();
    }

    private void AppendParagraphStyle(StringBuilder builder, UsedStyle used)
    {
        var formatting = document.Styles.Paragraph(used.StyleId);
        var declarations = new StringBuilder();
        AppendFormatting(declarations, formatting, includeRun: true);
        if (declarations.Length == 0)
            return;

        builder.Append('.').Append(used.ClassName).Append(" {\n");
        builder.Append(declarations);
        builder.Append("}\n");
    }

    private void AppendTableStyle(StringBuilder builder, UsedStyle used)
    {
        var info = document.Styles.Table(used.StyleId);
        builder.Append('.').Append(used.ClassName).Append(" {\n");
        builder.Append("  border-collapse: collapse;\n");
        if (info is not null)
            AppendFormatting(builder, info.Formatting, includeRun: true);
        builder.Append("}\n");
        if (info?.HasBorders == true)
        {
            builder.Append('.').Append(used.ClassName).Append(" td, .")
                .Append(used.ClassName).Append(" th {\n");
            builder.Append("  border: 1px solid #666666;\n");
            builder.Append("}\n");
        }
    }

    internal static void AppendFormatting(StringBuilder builder, ParagraphFormattingPatch formatting, bool includeRun)
    {
        if (includeRun)
            AppendRun(builder, formatting.Run);

        if (formatting.Alignment is string alignment)
            builder.Append("  text-align: ").Append(alignment).Append(";\n");

        if (formatting.SpacingBeforeTwips is int before)
            builder.Append("  margin-top: ").Append(Points(before)).Append(";\n");

        if (formatting.SpacingAfterTwips is int after)
            builder.Append("  margin-bottom: ").Append(Points(after)).Append(";\n");

        if (formatting.IndentLeftTwips is int left)
            builder.Append("  margin-left: ").Append(Points(left)).Append(";\n");

        if (formatting.IndentRightTwips is int right)
            builder.Append("  margin-right: ").Append(Points(right)).Append(";\n");

        if (formatting.FirstLineTwips is int firstLine)
            builder.Append("  text-indent: ").Append(Points(firstLine)).Append(";\n");
        else if (formatting.HangingTwips is int hanging)
            builder.Append("  text-indent: -").Append(Points(hanging)).Append(";\n");

        if (formatting.LineValue is int line)
        {
            if (formatting.LineRule is "exact" or "atleast")
                builder.Append("  line-height: ").Append(Points(line)).Append(";\n");
            else
                builder.Append("  line-height: ").Append(CssNumbers.Decimal(line / (double)UnitConverter.LineUnitPerSingleSpacing)).Append(";\n");
        }
    }

    private static void AppendRun(StringBuilder builder, RunFormattingPatch run)
    {
        if (!string.IsNullOrEmpty(run.FontFamily))
            builder.Append("  font-family: ").Append(QuoteFont(run.FontFamily)).Append(";\n");

        if (run.FontSizeHalfPoints is int halfPoints)
            builder.Append("  font-size: ").Append(CssNumbers.Points(UnitConverter.HalfPointsToPoints(halfPoints))).Append("pt;\n");

        if (run.Bold == true)
            builder.Append("  font-weight: bold;\n");
        else if (run.Bold == false)
            builder.Append("  font-weight: normal;\n");

        if (run.Italic == true)
            builder.Append("  font-style: italic;\n");
        else if (run.Italic == false)
            builder.Append("  font-style: normal;\n");

        if (!string.IsNullOrEmpty(run.ColorHex))
            builder.Append("  color: #").Append(run.ColorHex.ToLowerInvariant()).Append(";\n");

        if (!string.IsNullOrEmpty(run.HighlightHex))
            builder.Append("  background-color: #").Append(run.HighlightHex.ToLowerInvariant()).Append(";\n");

        var decorations = new List<string>();
        if (run.Underline == true)
            decorations.Add("underline");
        if (run.Strike == true)
            decorations.Add("line-through");
        if (decorations.Count > 0)
            builder.Append("  text-decoration: ").Append(string.Join(' ', decorations)).Append(";\n");

        if (run.VerticalPosition == VerticalTextPosition.Superscript)
            builder.Append("  vertical-align: super;\n");
        else if (run.VerticalPosition == VerticalTextPosition.Subscript)
            builder.Append("  vertical-align: sub;\n");
    }

    private string BaseRules()
    {
        var prefix = options.ClassPrefix ?? "";
        var fontFamily = string.IsNullOrWhiteSpace(document.BodyFontFamily) ? "Calibri" : document.BodyFontFamily;
        var fontSize = document.BodyFontSizeHalfPoints is int halfPoints
            ? CssNumbers.Points(UnitConverter.HalfPointsToPoints(halfPoints)) + "pt"
            : "11pt";
        var before = document.ParagraphSpacingBeforeTwips is int beforeTwips
            ? CssNumbers.Points(UnitConverter.TwipsToPoints(beforeTwips)) + "pt"
            : "0";
        var after = document.ParagraphSpacingAfterTwips is int afterTwips
            ? CssNumbers.Points(UnitConverter.TwipsToPoints(afterTwips)) + "pt"
            : "8pt";
        var lineHeight = "1.15";
        if (document.DefaultLineValue is int line)
        {
            lineHeight = document.DefaultLineRule is "exact" or "atleast"
                ? CssNumbers.Points(UnitConverter.TwipsToPoints(line)) + "pt"
                : CssNumbers.Decimal(line / (double)UnitConverter.LineUnitPerSingleSpacing);
        }

        var pageBreak = MarkupSafety.ClassToken(prefix + "page-break");
        var tab = MarkupSafety.ClassToken(prefix + "tab");
        var textbox = MarkupSafety.ClassToken(prefix + "textbox");
        var footnotes = MarkupSafety.ClassToken(prefix + "footnotes");
        var endnotes = MarkupSafety.ClassToken(prefix + "endnotes");
        var footnoteReference = MarkupSafety.ClassToken(prefix + "footnote-ref");
        var header = MarkupSafety.ClassToken(prefix + "doc-header");
        var footer = MarkupSafety.ClassToken(prefix + "doc-footer");
        var comment = MarkupSafety.ClassToken(prefix + "comment");
        var firstRow = MarkupSafety.ClassToken(prefix + "t-first-row");
        var banded = MarkupSafety.ClassToken(prefix + "t-banded");

        var builder = new StringBuilder();
        builder.Append("body {\n");
        builder.Append("  font-family: ").Append(QuoteFont(fontFamily)).Append(", Carlito, \"Segoe UI\", sans-serif;\n");
        builder.Append("  font-size: ").Append(fontSize).Append(";\n");
        builder.Append("  line-height: ").Append(lineHeight).Append(";\n");
        builder.Append("  color: #222222;\n");
        builder.Append("}\n");
        builder.Append("h1, h2, h3, h4, h5, h6 {\n");
        builder.Append("  margin: 0.8em 0 0.4em;\n");
        builder.Append("  font-weight: bold;\n");
        builder.Append("  line-height: 1.2;\n");
        builder.Append("}\n");
        builder.Append("p {\n");
        builder.Append("  margin: ").Append(before).Append(" 0 ").Append(after).Append(";\n");
        builder.Append("}\n");
        builder.Append("table {\n");
        builder.Append("  border-collapse: collapse;\n");
        builder.Append("  width: 100%;\n");
        builder.Append("  margin: 0 0 1em;\n");
        builder.Append("}\n");
        builder.Append("td, th {\n");
        builder.Append("  border: 1px solid #bbbbbb;\n");
        builder.Append("  padding: 4pt 6pt;\n");
        builder.Append("  vertical-align: top;\n");
        builder.Append("}\n");
        builder.Append("th {\n");
        builder.Append("  font-weight: bold;\n");
        builder.Append("  text-align: left;\n");
        builder.Append("}\n");
        builder.Append('.').Append(pageBreak).Append(" {\n  border: 0;\n  border-top: 1px solid #cccccc;\n  margin: 1.5em 0;\n}\n");
        builder.Append('.').Append(tab).Append(" {\n  display: inline-block;\n  width: 2em;\n}\n");
        builder.Append('.').Append(textbox).Append(" {\n  border: 1px solid #cccccc;\n  padding: 0.5em 0.75em;\n  margin: 0.5em 0;\n}\n");
        builder.Append('.').Append(footnotes).Append(", .").Append(endnotes).Append(" {\n  margin-top: 2em;\n  border-top: 1px solid #cccccc;\n  font-size: 0.92em;\n}\n");
        builder.Append('.').Append(footnotes).Append(" ol, .").Append(endnotes).Append(" ol {\n  padding-left: 1.5em;\n}\n");
        builder.Append('.').Append(footnoteReference).Append(" {\n  text-decoration: none;\n}\n");
        builder.Append('.').Append(header).Append(", .").Append(footer).Append(" {\n  color: #444444;\n}\n");
        builder.Append('.').Append(header).Append(" {\n  margin-bottom: 1.5em;\n}\n");
        builder.Append('.').Append(footer).Append(" {\n  margin-top: 1.5em;\n}\n");
        builder.Append('.').Append(comment).Append(" {\n  color: #666666;\n  font-size: 0.9em;\n}\n");
        AppendListStyle(builder, prefix, "lower-alpha");
        AppendListStyle(builder, prefix, "upper-alpha");
        AppendListStyle(builder, prefix, "lower-roman");
        AppendListStyle(builder, prefix, "upper-roman");
        AppendListStyle(builder, prefix, "decimal-leading-zero");
        builder.Append("table.").Append(firstRow).Append(" thead th {\n  background: #f3f6fb;\n}\n");
        builder.Append("table.").Append(banded).Append(" tbody tr:nth-child(even) {\n  background: #f7f7f7;\n}\n");
        builder.Append("img {\n  max-width: 100%;\n  height: auto;\n}\n");
        builder.Append("a {\n  color: #0563c1;\n}\n");
        return builder.ToString();
    }

    private static void AppendListStyle(StringBuilder builder, string prefix, string format)
    {
        var className = MarkupSafety.ClassToken(prefix + format);
        builder.Append("ol.").Append(className).Append(" {\n  list-style-type: ").Append(format).Append(";\n}\n");
    }

    private static string Points(int twips) =>
        CssNumbers.Points(UnitConverter.TwipsToPoints(twips)) + "pt";

    internal static string QuoteFont(string fontFamily)
    {
        var escaped = fontFamily.Replace("\\", "\\\\", StringComparison.Ordinal).Replace("\"", "\\\"", StringComparison.Ordinal);
        return "\"" + escaped + "\"";
    }
}
