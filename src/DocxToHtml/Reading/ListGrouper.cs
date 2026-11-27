using DocxToHtml.Internal;
using DocxToHtml.Model;

namespace DocxToHtml.Reading;

internal static class ListGrouper
{
    public static List<Block> Group(IReadOnlyList<Block> blocks, NumberingRepository numbering, WarningSink warnings)
    {
        var output = new List<Block>();
        var stack = new List<ListBlock>();

        foreach (var block in blocks)
        {
            if (block is ParagraphBlock paragraph && paragraph.List is not null)
            {
                AppendItem(paragraph);
                continue;
            }

            CloseAll();
            switch (block)
            {
                case TableBlock table:
                    foreach (var row in table.Rows)
                    {
                        foreach (var cell in row.Cells)
                            cell.Blocks = Group(cell.Blocks, numbering, warnings);
                    }
                    output.Add(table);
                    break;
                case TextBoxBlock textBox:
                    textBox.Blocks = Group(textBox.Blocks, numbering, warnings);
                    output.Add(textBox);
                    break;
                case ParagraphBlock plain:
                    foreach (var inline in plain.Inlines)
                    {
                        if (inline is TextBoxInline textBoxInline)
                            textBoxInline.Blocks = Group(textBoxInline.Blocks, numbering, warnings);
                    }
                    output.Add(plain);
                    break;
                default:
                    output.Add(block);
                    break;
            }
        }

        CloseAll();
        return output;

        void AppendItem(ParagraphBlock paragraph)
        {
            var listInformation = paragraph.List!;
            while (stack.Count > 0
                   && (stack[^1].NumberingId != listInformation.NumberingId || stack[^1].Level > listInformation.Level))
            {
                Pop();
            }

            if (stack.Count == 0
                || stack[^1].NumberingId != listInformation.NumberingId
                || stack[^1].Level < listInformation.Level)
            {
                var parentLevel = stack.Count == 0 ? -1 : stack[^1].Level;
                for (var level = parentLevel + 1; level <= listInformation.Level; level++)
                {
                    var described = numbering.Describe(listInformation.NumberingId, level, warnings, "list");
                    var opened = new ListBlock
                    {
                        NumberingId = listInformation.NumberingId,
                        Level = level,
                        IsBullet = described?.IsBullet ?? listInformation.IsBullet,
                        FormatClass = described?.FormatClass ?? listInformation.FormatClass,
                        Start = described?.Start ?? listInformation.Start
                    };

                    if (stack.Count > 0 && stack[^1].Items.Count == 0)
                        stack[^1].Items.Add(new ListItemModel());

                    stack.Add(opened);
                    if (level < listInformation.Level)
                        opened.Items.Add(new ListItemModel());
                }
            }

            stack[^1].Items.Add(new ListItemModel { Paragraph = paragraph });
        }

        void CloseAll()
        {
            while (stack.Count > 0)
                Pop();
        }

        void Pop()
        {
            var finished = stack[^1];
            stack.RemoveAt(stack.Count - 1);
            if (stack.Count == 0)
                output.Add(finished);
            else
                stack[^1].Items[^1].NestedLists.Add(finished);
        }
    }
}
