using DocxToHtml.Model;

namespace DocxToHtml.Reading;

internal static class RunMerger
{
    public static List<Block> MergeBlocks(List<Block> blocks)
    {
        foreach (var block in blocks)
            MergeBlock(block);

        return blocks;
    }

    public static List<Inline> Merge(List<Inline> inlines)
    {
        var merged = new List<Inline>();
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case TextRunInline textRun when merged.LastOrDefault() is TextRunInline previous
                    && previous.Formatting.VisuallyEquals(textRun.Formatting):
                    merged[^1] = previous with { Text = previous.Text + textRun.Text };
                    break;
                case HyperlinkInline hyperlink:
                    var mergedChildren = Merge(hyperlink.Children);
                    if (merged.LastOrDefault() is HyperlinkInline previousLink
                        && string.Equals(previousLink.Href, hyperlink.Href, StringComparison.Ordinal))
                    {
                        previousLink.Children.AddRange(mergedChildren);
                        previousLink.Children = Merge(previousLink.Children);
                    }
                    else
                    {
                        hyperlink.Children = mergedChildren;
                        merged.Add(hyperlink);
                    }
                    break;
                case TextBoxInline textBox:
                    textBox.Blocks = MergeBlocks(textBox.Blocks);
                    merged.Add(textBox);
                    break;
                default:
                    merged.Add(inline);
                    break;
            }
        }

        return merged;
    }

    private static void MergeBlock(Block block)
    {
        switch (block)
        {
            case ParagraphBlock paragraph:
                paragraph.Inlines = Merge(paragraph.Inlines);
                break;
            case TableBlock table:
                foreach (var row in table.Rows)
                {
                    foreach (var cell in row.Cells)
                        cell.Blocks = MergeBlocks(cell.Blocks);
                }
                break;
            case TextBoxBlock textBox:
                textBox.Blocks = MergeBlocks(textBox.Blocks);
                break;
            case ListBlock list:
                foreach (var item in list.Items)
                {
                    if (item.Paragraph is not null)
                        MergeBlock(item.Paragraph);

                    foreach (var nested in item.NestedLists)
                        MergeBlock(nested);
                }
                break;
        }
    }
}
