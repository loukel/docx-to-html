using DocumentFormat.OpenXml.Packaging;
using DocxToHtml.Internal;

namespace DocxToHtml.Reading;

internal sealed class NoteRequest
{
    public required int Identifier { get; init; }
    public required int DisplayNumber { get; init; }
    public required bool IsEndnote { get; init; }
}

internal sealed class ReadingContext
{
    public required DocxToHtmlOptions Options { get; init; }
    public required WarningSink Warnings { get; init; }
    public required StyleRepository Styles { get; init; }
    public required NumberingRepository Numbering { get; init; }
    public required ThemeColorMap ThemeColors { get; init; }
    public required ImageCollector Images { get; init; }
    public required CommentLookup Comments { get; init; }
    public FieldController Fields { get; } = new();
    public OpenXmlPart CurrentPart { get; set; } = null!;
    public string Location { get; set; } = "body";
    public List<NoteRequest> Notes { get; } = new();

    private readonly Dictionary<string, int> noteNumbers = new(StringComparer.Ordinal);
    private int nextFootnoteNumber = 1;
    private int nextEndnoteNumber = 1;

    public int AssignNoteNumber(int identifier, bool isEndnote)
    {
        var key = (isEndnote ? "e" : "f") + identifier.ToString(CultureInfo.InvariantCulture);
        if (noteNumbers.TryGetValue(key, out var existing))
            return existing;

        var displayNumber = isEndnote ? nextEndnoteNumber++ : nextFootnoteNumber++;
        noteNumbers[key] = displayNumber;
        Notes.Add(new NoteRequest
        {
            Identifier = identifier,
            DisplayNumber = displayNumber,
            IsEndnote = isEndnote
        });
        return displayNumber;
    }
}
