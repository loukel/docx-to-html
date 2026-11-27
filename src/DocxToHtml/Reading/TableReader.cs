using DocumentFormat.OpenXml;
using DocxToHtml.Internal;
using DocxToHtml.Model;

namespace DocxToHtml.Reading;

internal sealed class TableReader
{
    private readonly ReadingContext context;
    private readonly BlockReader blockReader;

    public TableReader(ReadingContext context, BlockReader blockReader)
    {
        this.context = context;
        this.blockReader = blockReader;
    }

    public TableBlock Read(OpenXmlElement table)
    {
        var tableProperties = OpenXmlValues.Child(table, "tblPr");
        var styleId = OpenXmlValues.Attribute(OpenXmlValues.Child(tableProperties, "tblStyle"), "val");
        var style = context.Styles.Find(styleId);
        var look = OpenXmlValues.Child(tableProperties, "tblLook");
        var rawRows = new List<RawRow>();

        foreach (var child in MarkupCompatibility.Children(table))
            CollectRows(child, rawRows);

        ApplyVerticalMerges(rawRows);

        var rows = new List<TableRowModel>();
        foreach (var rawRow in rawRows)
        {
            var row = new TableRowModel { IsHeader = rawRow.IsHeader };
            if (rawRow.GridBefore > 0)
                row.Cells.Add(new TableCellModel { ColumnSpan = rawRow.GridBefore });

            foreach (var cell in rawRow.Cells)
            {
                if (cell.Suppress)
                    continue;

                row.Cells.Add(new TableCellModel
                {
                    ColumnSpan = cell.ColumnSpan,
                    RowSpan = cell.RowSpan,
                    ShadingHex = cell.ShadingHex,
                    HorizontalAlignment = cell.HorizontalAlignment,
                    VerticalAlignment = cell.VerticalAlignment,
                    Blocks = cell.Blocks
                });
            }

            if (rawRow.GridAfter > 0)
                row.Cells.Add(new TableCellModel { ColumnSpan = rawRow.GridAfter });

            rows.Add(row);
        }

        return new TableBlock
        {
            StyleId = styleId,
            StyleName = style?.Name ?? styleId,
            FirstRowLook = OpenXmlValues.OnOffAttribute(look, "firstRow") == true,
            BandedRows = OpenXmlValues.OnOffAttribute(look, "noHBand") == false,
            ColumnWidthPercentages = ReadColumnWidths(table),
            Rows = rows
        };
    }

    private void CollectRows(OpenXmlElement element, List<RawRow> rows)
    {
        switch (element.LocalName)
        {
            case "tr":
                rows.Add(ReadRow(element));
                return;
            case "sdt":
                var content = OpenXmlValues.Child(element, "sdtContent");
                if (content is not null)
                {
                    foreach (var child in MarkupCompatibility.Children(content))
                        CollectRows(child, rows);
                }
                return;
            case "ins":
            case "del":
            case "moveFrom":
            case "moveTo":
                if (!RevisionFilter.Include(element.LocalName, context.Options.TrackedChanges))
                    return;
                foreach (var child in MarkupCompatibility.Children(element))
                    CollectRows(child, rows);
                return;
            case "customXml":
                foreach (var child in MarkupCompatibility.Children(element))
                    CollectRows(child, rows);
                return;
        }
    }

    private RawRow ReadRow(OpenXmlElement row)
    {
        var rowProperties = OpenXmlValues.Child(row, "trPr");
        var header = OpenXmlValues.Child(rowProperties, "tblHeader");
        var rawRow = new RawRow
        {
            IsHeader = header is not null && OpenXmlValues.OnOffAttribute(header, "val") != false,
            GridBefore = Math.Max(0, OpenXmlValues.IntegerAttribute(OpenXmlValues.Child(rowProperties, "gridBefore"), "val") ?? 0),
            GridAfter = Math.Max(0, OpenXmlValues.IntegerAttribute(OpenXmlValues.Child(rowProperties, "gridAfter"), "val") ?? 0)
        };

        var columnIndex = rawRow.GridBefore;
        foreach (var child in MarkupCompatibility.Children(row))
        {
            if (child.LocalName == "sdt")
            {
                var content = OpenXmlValues.Child(child, "sdtContent");
                if (content is null)
                    continue;

                foreach (var nested in MarkupCompatibility.Children(content))
                {
                    if (nested.LocalName == "tc")
                        columnIndex = AddCell(rawRow, nested, columnIndex);
                }
                continue;
            }

            if (child.LocalName == "tc")
                columnIndex = AddCell(rawRow, child, columnIndex);
        }

        return rawRow;
    }

    private int AddCell(RawRow row, OpenXmlElement cellElement, int columnIndex)
    {
        var cellProperties = OpenXmlValues.Child(cellElement, "tcPr");
        var columnSpan = OpenXmlValues.IntegerAttribute(OpenXmlValues.Child(cellProperties, "gridSpan"), "val") ?? 1;
        if (columnSpan < 1)
            columnSpan = 1;

        var mergeElement = OpenXmlValues.Child(cellProperties, "vMerge");
        var merge = VerticalMergeKind.None;
        if (mergeElement is not null)
        {
            var value = OpenXmlValues.Attribute(mergeElement, "val");
            merge = value is not null && value.Equals("restart", StringComparison.OrdinalIgnoreCase)
                ? VerticalMergeKind.Restart
                : VerticalMergeKind.Continue;
        }

        var fill = OpenXmlValues.Attribute(OpenXmlValues.Child(cellProperties, "shd"), "fill");
        string? shading = fill is { Length: 6 } && !fill.Equals("auto", StringComparison.OrdinalIgnoreCase)
            ? fill.ToUpperInvariant()
            : null;

        row.Cells.Add(new RawCell
        {
            ColumnIndex = columnIndex,
            ColumnSpan = columnSpan,
            Merge = merge,
            ShadingHex = shading,
            VerticalAlignment = ReadVerticalAlignment(cellProperties),
            HorizontalAlignment = ReadHorizontalAlignment(cellElement),
            Blocks = blockReader.ReadBlocks(cellElement)
        });
        return columnIndex + columnSpan;
    }

    private string? ReadHorizontalAlignment(OpenXmlElement cellElement)
    {
        foreach (var child in MarkupCompatibility.Children(cellElement))
        {
            if (child.LocalName != "p")
                continue;

            return PropertyReader.ReadParagraph(
                OpenXmlValues.Child(child, "pPr"),
                context.ThemeColors,
                context.Warnings,
                context.Location).Alignment;
        }

        return null;
    }

    private static string? ReadVerticalAlignment(OpenXmlElement? cellProperties)
    {
        return OpenXmlValues.Attribute(OpenXmlValues.Child(cellProperties, "vAlign"), "val")?.ToLowerInvariant() switch
        {
            "center" => "middle",
            "bottom" => "bottom",
            "top" => "top",
            _ => null
        };
    }

    private void ApplyVerticalMerges(List<RawRow> rows)
    {
        var activeMerges = new List<ActiveMerge>();
        foreach (var row in rows)
        {
            foreach (var merge in activeMerges)
                merge.Continued = false;

            var started = new List<ActiveMerge>();
            foreach (var cell in row.Cells)
            {
                var match = activeMerges.FirstOrDefault(candidate =>
                    cell.ColumnIndex >= candidate.StartColumn && cell.ColumnIndex < candidate.EndColumn);

                if (cell.Merge == VerticalMergeKind.Continue && match is not null)
                {
                    cell.Suppress = true;
                    match.Continued = true;
                    if (cell.Blocks.Any(HasVisibleContent))
                    {
                        match.Cell.Blocks.AddRange(cell.Blocks);
                        context.Warnings.Add(
                            "MergedCellContent",
                            "Text in a vertically merged continuation cell was kept in the starting cell.",
                            context.Location);
                    }
                }
                else if (cell.Merge == VerticalMergeKind.Restart)
                {
                    started.Add(new ActiveMerge(cell, cell.ColumnIndex, cell.ColumnIndex + cell.ColumnSpan));
                }
                else if (cell.Merge == VerticalMergeKind.Continue)
                {
                    context.Warnings.Add(
                        "OrphanedMerge",
                        "A vertical merge continuation had no starting cell and was written as its own cell.",
                        context.Location);
                }
            }

            foreach (var merge in activeMerges)
            {
                if (merge.Continued)
                    merge.Cell.RowSpan++;
            }

            activeMerges = activeMerges.Where(merge => merge.Continued).Concat(started).ToList();
        }
    }

    private static bool HasVisibleContent(Block block)
    {
        return block switch
        {
            ParagraphBlock paragraph => paragraph.HasVisibleContent(),
            TableBlock table => table.Rows.Count > 0,
            TextBoxBlock textBox => textBox.Blocks.Count > 0,
            ListBlock list => list.Items.Count > 0,
            _ => true
        };
    }

    private static List<int> ReadColumnWidths(OpenXmlElement table)
    {
        var grid = OpenXmlValues.Child(table, "tblGrid");
        if (grid is null)
            return new List<int>();

        var widths = new List<int>();
        foreach (var column in MarkupCompatibility.Children(grid))
        {
            if (column.LocalName != "gridCol")
                continue;

            var width = OpenXmlValues.IntegerAttribute(column, "w");
            if (width is > 0)
                widths.Add(width.Value);
        }

        if (widths.Count == 0 || widths.Sum() <= 0)
            return new List<int>();

        var total = widths.Sum();
        var percentages = widths
            .Select(width => (int)Math.Round(width * 100d / total, MidpointRounding.AwayFromZero))
            .ToList();
        var drift = 100 - percentages.Sum();
        percentages[^1] = Math.Max(1, percentages[^1] + drift);
        return percentages;
    }

    private enum VerticalMergeKind
    {
        None,
        Restart,
        Continue
    }

    private sealed class RawRow
    {
        public bool IsHeader { get; init; }
        public int GridBefore { get; init; }
        public int GridAfter { get; init; }
        public List<RawCell> Cells { get; } = new();
    }

    private sealed class RawCell
    {
        public int ColumnIndex { get; init; }
        public int ColumnSpan { get; init; }
        public int RowSpan { get; set; } = 1;
        public VerticalMergeKind Merge { get; init; }
        public bool Suppress { get; set; }
        public string? ShadingHex { get; init; }
        public string? HorizontalAlignment { get; init; }
        public string? VerticalAlignment { get; init; }
        public List<Block> Blocks { get; init; } = new();
    }

    private sealed class ActiveMerge
    {
        public ActiveMerge(RawCell cell, int startColumn, int endColumn)
        {
            Cell = cell;
            StartColumn = startColumn;
            EndColumn = endColumn;
        }

        public RawCell Cell { get; }
        public int StartColumn { get; }
        public int EndColumn { get; }
        public bool Continued { get; set; }
    }
}
