namespace DocxToHtml.Writing;

internal sealed class UsedStyle
{
    public required string ClassName { get; init; }
    public required string StyleId { get; init; }
    public bool IsTable { get; init; }
}
