using DocxToHtml;

return CliApplication.Run(args);

internal static class CliApplication
{
    public static int Run(string[] arguments)
    {
        if (arguments.Length < 2 || arguments.Contains("--help", StringComparer.Ordinal))
        {
            Console.Error.WriteLine("Usage: DocxToHtml.Cli <input.docx> <output-directory> [--prefix dx-] [--fragment] [--embed-css] [--headers] [--comments] [--reject-changes] [--images external|data-uri|skip] [--title \"Document title\"]");
            return arguments.Contains("--help", StringComparer.Ordinal) ? 0 : 1;
        }

        var options = new DocxToHtmlOptions();
        for (var index = 2; index < arguments.Length; index++)
        {
            var argument = arguments[index];
            switch (argument)
            {
                case "--fragment":
                    options.OutputMode = HtmlOutputMode.BodyFragment;
                    break;
                case "--embed-css":
                    options.CssMode = CssMode.Embedded;
                    break;
                case "--headers":
                    options.IncludeHeadersAndFooters = true;
                    break;
                case "--comments":
                    options.IncludeComments = true;
                    break;
                case "--reject-changes":
                    options.TrackedChanges = TrackedChangesMode.Reject;
                    break;
                case "--prefix":
                    options.ClassPrefix = RequireValue(arguments, ref index, argument);
                    break;
                case "--title":
                    options.DocumentTitle = RequireValue(arguments, ref index, argument);
                    break;
                case "--images":
                    options.ImageMode = RequireValue(arguments, ref index, argument).ToLowerInvariant() switch
                    {
                        "external" => ImageMode.External,
                        "data-uri" => ImageMode.DataUri,
                        "skip" => ImageMode.Skip,
                        var other => throw new ArgumentException($"Unknown image mode '{other}'.")
                    };
                    break;
                default:
                    throw new ArgumentException($"Unknown argument '{argument}'.");
            }
        }

        var result = DocxToHtmlConverter.ConvertToFolder(arguments[0], arguments[1], options);
        Console.WriteLine($"Wrote index.html ({result.Html.Length} characters), {result.Images.Count} image(s), {result.Headings.Count} heading(s), {result.Warnings.Count} warning(s).");
        foreach (var warning in result.Warnings)
            Console.WriteLine($"warning {warning.Code}: {warning.Message}");

        return 0;
    }

    private static string RequireValue(string[] arguments, ref int index, string name)
    {
        if (index + 1 >= arguments.Length)
            throw new ArgumentException($"Missing value for {name}.");

        index++;
        return arguments[index];
    }
}
