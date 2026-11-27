namespace DocxToHtml.Model;

internal abstract class Block;

internal sealed class ParagraphBlock : Block
{
    public string? StyleId { get; init; }
    public string? StyleName { get; init; }
    public bool IsNormalStyle { get; init; }
    public ParagraphKind Kind { get; init; }
    public int? HeadingLevel { get; init; }
    public string? MappedTag { get; init; }
    public string? MappedClass { get; init; }
    public string? ElementId { get; set; }
    public List<Inline> Inlines { get; set; } = new();
    public ListInformation? List { get; init; }
    public RunFormatting Baseline { get; init; } = RunFormatting.Empty;
    public ParagraphFormattingPatch DirectFormatting { get; init; } = ParagraphFormattingPatch.Empty;

    public bool HasVisibleContent()
    {
        if (!string.IsNullOrEmpty(ElementId))
            return true;

        return Inlines.Any(IsVisibleInline);
    }

    public static bool IsVisibleInline(Inline inline)
    {
        return inline switch
        {
            TextRunInline textRun => textRun.Text.Length > 0,
            BookmarkInline => false,
            _ => true
        };
    }
}

internal sealed class TableBlock : Block
{
    public string? StyleId { get; init; }
    public string? StyleName { get; init; }
    public bool FirstRowLook { get; init; }
    public bool BandedRows { get; init; }
    public List<int> ColumnWidthPercentages { get; init; } = new();
    public List<TableRowModel> Rows { get; init; } = new();
}

internal sealed class TableRowModel
{
    public bool IsHeader { get; init; }
    public List<TableCellModel> Cells { get; init; } = new();
}

internal sealed class TableCellModel
{
    public int ColumnSpan { get; init; } = 1;
    public int RowSpan { get; init; } = 1;
    public string? ShadingHex { get; init; }
    public string? HorizontalAlignment { get; init; }
    public string? VerticalAlignment { get; init; }
    public List<Block> Blocks { get; set; } = new();
}

internal sealed class TextBoxBlock : Block
{
    public List<Block> Blocks { get; set; } = new();
}

internal sealed class ListBlock : Block
{
    public int NumberingId { get; init; }
    public int Level { get; init; }
    public bool IsBullet { get; init; }
    public string? FormatClass { get; init; }
    public int? Start { get; init; }
    public List<ListItemModel> Items { get; init; } = new();
}

internal sealed class ListItemModel
{
    public ParagraphBlock? Paragraph { get; init; }
    public List<ListBlock> NestedLists { get; init; } = new();
}

internal abstract record Inline;

internal sealed record TextRunInline(string Text, RunFormatting Formatting) : Inline;

internal sealed record LineBreakInline : Inline;

internal sealed record PageBreakInline : Inline;

internal sealed record TabInline : Inline;

internal sealed record NonBreakingHyphenInline : Inline;

internal sealed record SoftHyphenInline : Inline;

internal sealed record SymbolInline(string Character) : Inline;

internal sealed record BookmarkInline(string Identifier) : Inline;

internal sealed record NoteReferenceInline(int DisplayNumber, bool IsEndnote) : Inline;

internal sealed record CommentInline(string Text) : Inline;

internal sealed record HyperlinkInline : Inline
{
    public required string Href { get; init; }
    public List<Inline> Children { get; set; } = new();
}

internal sealed record ImageInline : Inline
{
    public required string FileName { get; init; }
    public required string ContentType { get; init; }
    public required byte[] Data { get; init; }
    public string AltText { get; init; } = "";
    public int? WidthPixels { get; init; }
    public int? HeightPixels { get; init; }
    public bool SourceIsUri { get; init; }
}

internal sealed record TextBoxInline : Inline
{
    public List<Block> Blocks { get; set; } = new();
}

internal sealed class NoteModel
{
    public int DisplayNumber { get; init; }
    public bool IsEndnote { get; init; }
    public List<Block> Blocks { get; set; } = new();
}

internal sealed class StyleCatalog
{
    private readonly Dictionary<string, ParagraphFormattingPatch> paragraphStyles = new(StringComparer.Ordinal);
    private readonly Dictionary<string, TableStyleInfo> tableStyles = new(StringComparer.Ordinal);

    public void AddParagraph(string styleId, ParagraphFormattingPatch formatting) =>
        paragraphStyles[styleId] = formatting;

    public ParagraphFormattingPatch Paragraph(string? styleId) =>
        styleId is not null && paragraphStyles.TryGetValue(styleId, out var formatting)
            ? formatting
            : ParagraphFormattingPatch.Empty;

    public void AddTable(string styleId, TableStyleInfo info) =>
        tableStyles[styleId] = info;

    public TableStyleInfo? Table(string? styleId) =>
        styleId is not null && tableStyles.TryGetValue(styleId, out var info) ? info : null;
}

internal sealed class DocumentModel
{
    public required string Title { get; init; }
    public string Language { get; init; } = "en";
    public List<Block> BodyBlocks { get; init; } = new();
    public List<Block> HeaderBlocks { get; init; } = new();
    public List<Block> FooterBlocks { get; init; } = new();
    public List<NoteModel> Notes { get; init; } = new();
    public List<ExtractedImage> Images { get; init; } = new();
    public IReadOnlyList<ConversionWarning> Warnings { get; init; } = Array.Empty<ConversionWarning>();
    public StyleCatalog Styles { get; init; } = new();
    public string? BodyFontFamily { get; init; }
    public int? BodyFontSizeHalfPoints { get; init; }
    public int? ParagraphSpacingBeforeTwips { get; init; }
    public int? ParagraphSpacingAfterTwips { get; init; }
    public int? DefaultLineValue { get; init; }
    public string? DefaultLineRule { get; init; }
}
