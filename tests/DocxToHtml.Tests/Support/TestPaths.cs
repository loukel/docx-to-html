namespace DocxToHtml.Tests.Support;

internal static class TestPaths
{
    public static string ProjectDirectory()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "DocxToHtml.Tests.csproj")))
                return directory.FullName;

            directory = directory.Parent;
        }

        throw new InvalidOperationException("Could not locate the test project directory.");
    }
}
