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
        AppendFormatting(declarations, formatting, includeRun: true, DocumentBaseline());
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

    internal static void AppendFormatting(
        StringBuilder builder,
        ParagraphFormattingPatch formatting,
        bool includeRun,
        ParagraphFormattingPatch? baseline = null)
    {
        if (includeRun)
            AppendRun(builder, formatting.Run, baseline?.Run);

        if (formatting.Alignment is string alignment
            && !alignment.Equals("left", StringComparison.OrdinalIgnoreCase)
            && !SameText(alignment, baseline?.Alignment))
        {
            builder.Append("  text-align: ").Append(alignment).Append(";\n");
        }

        AppendBoxSide(builder, "margin-top", formatting.SpacingBeforeTwips, baseline?.SpacingBeforeTwips);
        AppendBoxSide(builder, "margin-bottom", formatting.SpacingAfterTwips, baseline?.SpacingAfterTwips);
        AppendBoxSide(builder, "margin-left", formatting.IndentLeftTwips, baseline?.IndentLeftTwips);
        AppendBoxSide(builder, "margin-right", formatting.IndentRightTwips, baseline?.IndentRightTwips);

        if (formatting.FirstLineTwips is int firstLine && !SameTwips(firstLine, baseline?.FirstLineTwips))
            builder.Append("  text-indent: ").Append(Points(firstLine)).Append(";\n");
        else if (formatting.HangingTwips is int hanging && !SameTwips(hanging, baseline?.HangingTwips))
            builder.Append("  text-indent: -").Append(Points(hanging)).Append(";\n");

        if (formatting.LineValue is int line
            && (baseline?.LineValue != line || !SameText(formatting.LineRule, baseline?.LineRule)))
        {
            if (formatting.LineRule is "exact" or "atleast")
                builder.Append("  line-height: ").Append(Points(line)).Append(";\n");
            else
                builder.Append("  line-height: ").Append(CssNumbers.Decimal(line / (double)UnitConverter.LineUnitPerSingleSpacing)).Append(";\n");
        }
    }

    private ParagraphFormattingPatch DocumentBaseline() => new()
    {
        SpacingBeforeTwips = document.ParagraphSpacingBeforeTwips,
        SpacingAfterTwips = document.ParagraphSpacingAfterTwips,
        LineValue = document.DefaultLineValue,
        LineRule = document.DefaultLineRule,
        Run = new RunFormattingPatch
        {
            FontFamily = document.BodyFontFamily,
            FontSizeHalfPoints = document.BodyFontSizeHalfPoints
        }
    };

    private static void AppendBoxSide(StringBuilder builder, string property, int? value, int? baseline)
    {
        if (value is not int twips || SameTwips(twips, baseline))
            return;

        builder.Append("  ").Append(property).Append(": ").Append(Points(twips)).Append(";\n");
    }

    private static bool SameTwips(int value, int? baseline) => value == (baseline ?? 0);

    private static bool SameText(string? value, string? baseline) =>
        string.Equals(value ?? "", baseline ?? "", StringComparison.OrdinalIgnoreCase);

    private static void AppendRun(StringBuilder builder, RunFormattingPatch run, RunFormattingPatch? baseline)
    {
        if (!string.IsNullOrEmpty(run.FontFamily) && !SameText(run.FontFamily, baseline?.FontFamily))
            builder.Append("  font-family: ").Append(QuoteFont(run.FontFamily)).Append(";\n");

        if (run.FontSizeHalfPoints is int halfPoints && halfPoints != baseline?.FontSizeHalfPoints)
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
        builder.Append("ul, ol {\n");
        builder.Append("  margin: 0 0 ").Append(after).Append(";\n");
        builder.Append("  padding-left: 1.5em;\n");
        builder.Append("}\n");
        if (Tables().Any())
        {
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
            if (Tables().Any(table => table.FirstRowLook))
                builder.Append("table.").Append(firstRow).Append(" thead th {\n  background: #f3f6fb;\n}\n");
            if (Tables().Any(table => table.BandedRows))
                builder.Append("table.").Append(banded).Append(" tbody tr:nth-child(even) {\n  background: #f7f7f7;\n}\n");
        }

        builder.Append("img {\n  max-width: 100%;\n  height: auto;\n}\n");
        builder.Append("a {\n  color: #0563c1;\n}\n");
        return builder.ToString();
    }

    private IEnumerable<TableBlock> Tables()
    {
        foreach (var block in EnumerateBlocks(document.BodyBlocks))
        {
            if (block is TableBlock table)
                yield return table;
        }

        foreach (var block in EnumerateBlocks(document.HeaderBlocks))
        {
            if (block is TableBlock table)
                yield return table;
        }

        foreach (var block in EnumerateBlocks(document.FooterBlocks))
        {
            if (block is TableBlock table)
                yield return table;
        }

        foreach (var note in document.Notes)
        {
            foreach (var block in EnumerateBlocks(note.Blocks))
            {
                if (block is TableBlock table)
                    yield return table;
            }
        }
    }

    private static IEnumerable<Block> EnumerateBlocks(IEnumerable<Block> blocks)
    {
        foreach (var block in blocks)
        {
            yield return block;
            IEnumerable<Block> nested = block switch
            {
                TableBlock table => table.Rows.SelectMany(row => row.Cells).SelectMany(cell => cell.Blocks),
                TextBoxBlock textBox => textBox.Blocks,
                ListBlock list => list.Items.SelectMany(item => item.NestedLists),
                ParagraphBlock paragraph => paragraph.Inlines.OfType<TextBoxInline>().SelectMany(textBox => textBox.Blocks),
                _ => Array.Empty<Block>()
            };

            foreach (var child in EnumerateBlocks(nested))
                yield return child;
        }
    }

    private static string Points(int twips) =>
        CssNumbers.Points(UnitConverter.TwipsToPoints(twips)) + "pt";

    internal static string QuoteFont(string fontFamily)
    {
        var escaped = fontFamily.Replace("\\", "\\\\", StringComparison.Ordinal).Replace("\"", "\\\"", StringComparison.Ordinal);
        return "\"" + escaped + "\"";
    }
}
