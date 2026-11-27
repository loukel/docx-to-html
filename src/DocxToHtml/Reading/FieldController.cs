namespace DocxToHtml.Reading;

internal sealed class FieldController
{
    private readonly Stack<FieldFrame> frames = new();

    public bool HasOpenField => frames.Count > 0;

    public bool ShouldSkipText => frames.Count > 0 && !frames.Peek().InResult;

    public string? CurrentHyperlinkHref
    {
        get
        {
            foreach (var frame in frames)
            {
                if (!frame.InResult)
                    return null;

                if (!string.IsNullOrEmpty(frame.Href))
                    return frame.Href;
            }

            return null;
        }
    }

    public void Begin() => frames.Push(new FieldFrame());

    public void AppendInstruction(string? text)
    {
        if (frames.Count == 0 || frames.Peek().InResult || string.IsNullOrEmpty(text))
            return;

        frames.Peek().Instruction.Append(text);
    }

    public void Separate()
    {
        if (frames.Count == 0)
            return;

        var frame = frames.Peek();
        frame.InResult = true;
        if (HyperlinkInstruction.TryParse(frame.Instruction.ToString(), out var href))
            frame.Href = href;
    }

    public void End()
    {
        if (frames.Count > 0)
            frames.Pop();
    }

    public void PushResult(string? href) =>
        frames.Push(new FieldFrame { InResult = true, Href = href });

    public void Pop()
    {
        if (frames.Count > 0)
            frames.Pop();
    }

    public void Clear() => frames.Clear();

    private sealed class FieldFrame
    {
        public StringBuilder Instruction { get; } = new();
        public bool InResult { get; set; }
        public string? Href { get; set; }
    }
}

internal static class HyperlinkInstruction
{
    public static bool TryParse(string? instruction, out string href)
    {
        href = "";
        if (string.IsNullOrWhiteSpace(instruction))
            return false;

        var trimmed = instruction.Trim();
        if (!trimmed.StartsWith("HYPERLINK", StringComparison.OrdinalIgnoreCase))
            return false;

        var quoted = ReadQuotedTarget(trimmed);
        if (quoted is null)
            return false;

        var localAnchor = System.Text.RegularExpressions.Regex.IsMatch(
            trimmed,
            @"\\l\b",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        href = localAnchor ? "#" + quoted : quoted;
        return href.Length > 0;
    }

    private static string? ReadQuotedTarget(string instruction)
    {
        var firstQuote = instruction.IndexOf('"');
        if (firstQuote >= 0)
        {
            var secondQuote = instruction.IndexOf('"', firstQuote + 1);
            if (secondQuote > firstQuote)
                return instruction[(firstQuote + 1)..secondQuote];
        }

        var tokens = instruction.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        return tokens.Length >= 2 ? tokens[1] : null;
    }
}
