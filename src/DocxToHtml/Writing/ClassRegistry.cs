using DocxToHtml.Internal;

namespace DocxToHtml.Writing;

internal sealed record RegisteredClass(string ClassName, IReadOnlyList<string> Declarations);

internal sealed class ClassRegistry
{
    private readonly string prefix;
    private readonly Dictionary<string, string> classNameBySignature = new(StringComparer.Ordinal);
    private readonly List<RegisteredClass> classes = new();

    public ClassRegistry(DocxToHtmlOptions options)
    {
        prefix = options.ClassPrefix ?? "";
    }

    public IReadOnlyList<RegisteredClass> Classes => classes;

    public string Register(string signature, params string[] declarations)
    {
        if (classNameBySignature.TryGetValue(signature, out var existing))
            return existing;

        var className = MarkupSafety.ClassToken(prefix + signature);
        classNameBySignature.Add(signature, className);
        classes.Add(new RegisteredClass(className, declarations));
        return className;
    }
}
