namespace DocxToHtml.Internal;

internal sealed class WarningSink
{
    private readonly List<ConversionWarning> warnings = new();
    private readonly Action<ConversionWarning>? onWarning;

    public WarningSink(Action<ConversionWarning>? onWarning)
    {
        this.onWarning = onWarning;
    }

    public IReadOnlyList<ConversionWarning> Warnings => warnings;

    public void Add(string code, string message, string? location = null)
    {
        var warning = new ConversionWarning(code, message, location);
        warnings.Add(warning);
        onWarning?.Invoke(warning);
    }
}
