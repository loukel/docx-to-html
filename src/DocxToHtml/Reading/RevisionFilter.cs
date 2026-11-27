namespace DocxToHtml.Reading;

internal static class RevisionFilter
{
    public static bool Include(string localName, TrackedChangesMode mode)
    {
        var acceptingChanges = mode != TrackedChangesMode.Reject;
        return localName switch
        {
            "ins" or "moveTo" => acceptingChanges,
            "del" or "moveFrom" => !acceptingChanges,
            _ => true
        };
    }
}
