namespace DocxToHtml.Model;

internal enum VerticalTextPosition
{
    Baseline,
    Superscript,
    Subscript
}

internal enum ParagraphKind
{
    Normal,
    Styled,
    Title,
    Subtitle,
    Mapped
}

internal sealed record RunFormatting
{
    public static RunFormatting Empty { get; } = new();

    public bool Bold { get; init; }
    public bool Italic { get; init; }
    public bool Underline { get; init; }
    public bool Strike { get; init; }
    public VerticalTextPosition VerticalPosition { get; init; }
    public string? ColorHex { get; init; }
    public string? HighlightHex { get; init; }
    public int? FontSizeHalfPoints { get; init; }
    public string? FontFamily { get; init; }

    public bool VisuallyEquals(RunFormatting other)
    {
        return Bold == other.Bold
            && Italic == other.Italic
            && Underline == other.Underline
            && Strike == other.Strike
            && VerticalPosition == other.VerticalPosition
            && string.Equals(ColorHex, other.ColorHex, StringComparison.OrdinalIgnoreCase)
            && string.Equals(HighlightHex, other.HighlightHex, StringComparison.OrdinalIgnoreCase)
            && FontSizeHalfPoints == other.FontSizeHalfPoints
            && string.Equals(FontFamily, other.FontFamily, StringComparison.OrdinalIgnoreCase);
    }
}

internal sealed record RunFormattingPatch
{
    public static RunFormattingPatch Empty { get; } = new();

    public bool? Bold { get; init; }
    public bool? Italic { get; init; }
    public bool? Underline { get; init; }
    public bool? Strike { get; init; }
    public VerticalTextPosition? VerticalPosition { get; init; }
    public string? ColorHex { get; init; }
    public string? HighlightHex { get; init; }
    public int? FontSizeHalfPoints { get; init; }
    public string? FontFamily { get; init; }

    public bool HasAnyValue =>
        Bold is not null
        || Italic is not null
        || Underline is not null
        || Strike is not null
        || VerticalPosition is not null
        || ColorHex is not null
        || HighlightHex is not null
        || FontSizeHalfPoints is not null
        || FontFamily is not null;

    public RunFormattingPatch Overlay(RunFormattingPatch upper) => new()
    {
        Bold = upper.Bold ?? Bold,
        Italic = upper.Italic ?? Italic,
        Underline = upper.Underline ?? Underline,
        Strike = upper.Strike ?? Strike,
        VerticalPosition = upper.VerticalPosition ?? VerticalPosition,
        ColorHex = upper.ColorHex ?? ColorHex,
        HighlightHex = upper.HighlightHex ?? HighlightHex,
        FontSizeHalfPoints = upper.FontSizeHalfPoints ?? FontSizeHalfPoints,
        FontFamily = upper.FontFamily ?? FontFamily
    };

    public RunFormatting Apply(RunFormatting baseline) => new()
    {
        Bold = Bold ?? baseline.Bold,
        Italic = Italic ?? baseline.Italic,
        Underline = Underline ?? baseline.Underline,
        Strike = Strike ?? baseline.Strike,
        VerticalPosition = VerticalPosition ?? baseline.VerticalPosition,
        ColorHex = ColorHex ?? baseline.ColorHex,
        HighlightHex = HighlightHex ?? baseline.HighlightHex,
        FontSizeHalfPoints = FontSizeHalfPoints ?? baseline.FontSizeHalfPoints,
        FontFamily = FontFamily ?? baseline.FontFamily
    };
}

internal sealed record ParagraphFormattingPatch
{
    public static ParagraphFormattingPatch Empty { get; } = new();

    public string? Alignment { get; init; }
    public int? SpacingBeforeTwips { get; init; }
    public int? SpacingAfterTwips { get; init; }
    public int? IndentLeftTwips { get; init; }
    public int? IndentRightTwips { get; init; }
    public int? FirstLineTwips { get; init; }
    public int? HangingTwips { get; init; }
    public int? LineValue { get; init; }
    public string? LineRule { get; init; }
    public RunFormattingPatch Run { get; init; } = RunFormattingPatch.Empty;

    public ParagraphFormattingPatch Overlay(ParagraphFormattingPatch upper) => new()
    {
        Alignment = upper.Alignment ?? Alignment,
        SpacingBeforeTwips = upper.SpacingBeforeTwips ?? SpacingBeforeTwips,
        SpacingAfterTwips = upper.SpacingAfterTwips ?? SpacingAfterTwips,
        IndentLeftTwips = upper.IndentLeftTwips ?? IndentLeftTwips,
        IndentRightTwips = upper.IndentRightTwips ?? IndentRightTwips,
        FirstLineTwips = upper.FirstLineTwips ?? FirstLineTwips,
        HangingTwips = upper.HangingTwips ?? HangingTwips,
        LineValue = upper.LineValue ?? LineValue,
        LineRule = upper.LineRule ?? LineRule,
        Run = Run.Overlay(upper.Run)
    };
}

internal abstract record NumberingSlot
{
    public static NumberingSlot Unset { get; } = new UnsetNumbering();

    internal sealed record UnsetNumbering : NumberingSlot;

    internal sealed record ClearedNumbering : NumberingSlot;

    internal sealed record ActiveNumbering(int NumberingId, int Level) : NumberingSlot;
}

internal sealed record ListInformation(int NumberingId, int Level, bool IsBullet, string? FormatClass, int? Start);

internal sealed record TableStyleInfo(ParagraphFormattingPatch Formatting, bool HasBorders);
