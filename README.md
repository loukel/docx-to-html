# DocxToHtml

DocxToHtml is a .NET 10 class library that converts a Word `.docx` file into semantic HTML and a separate stylesheet. The HTML never contains a `style` attribute. Presentation lives in CSS classes.

```csharp
using DocxToHtml;

var result = DocxToHtmlConverter.Convert("notes.docx");
File.WriteAllText("notes.html", result.Html);
File.WriteAllText("notes.css", result.Css);

foreach (var heading in result.Headings)
    Console.WriteLine($"{heading.Level}: {heading.Text}");
```

`result.Headings` lists every heading in document order. The level is the Word level, including 7, 8, and 9. HTML only has `h1`–`h6`, so a deeper heading is written as `h6` with an extra class `h7`, `h8`, or `h9`:

```html
<h6 class="p-heading-7 h7">Deep</h6>
```

With `ClassPrefix = "dx-"` that class is `dx-h7`. A subtitle is a paragraph, not a heading. The Title style is `h1` with class `title`.

## Projects

- `src/DocxToHtml` — the library. Public types are in the `DocxToHtml` namespace.
- `samples/DocxToHtml.Cli` — a small command that writes a folder.
- `tests/DocxToHtml.Tests` — one fixture per feature, snapshot HTML and CSS, and invariants on every fixture.

The only runtime dependency is DocumentFormat.OpenXml. The library does not write to disk or the console except `ConvertToFolder`.

## Requirements

.NET SDK 10.

```bash
dotnet build DocxToHtml.sln
dotnet test DocxToHtml.sln
```

## Convert

```csharp
ConversionResult DocxToHtmlConverter.Convert(string docxPath, DocxToHtmlOptions? options = null);
ConversionResult DocxToHtmlConverter.Convert(Stream docxStream, DocxToHtmlOptions? options = null);
ConversionResult DocxToHtmlConverter.ConvertToFolder(string docxPath, string outputDirectory, DocxToHtmlOptions? options = null);
```

`Convert(Stream)` copies the stream and does not dispose the caller's stream. A file that is not a package throws `InvalidDataException`. `ConvertToFolder` writes `index.html`, the stylesheet named by `CssFileName` (default `styles.css`), and image files under `ImageFolder` (default `images`). The HTML file is indented: each block sits on its own line, and the text of a heading, paragraph, or table cell stays on that element's line.

`ConversionResult` contains `Html`, `Css`, `Images`, `Warnings`, and `Headings`.

## Options

Create a `DocxToHtmlOptions` and pass it to any convert method. The converter does not modify the instance.

### OutputMode

`HtmlOutputMode.FullDocument` (default) returns a complete document. `BodyFragment` returns only the body contents.

```csharp
var fragment = DocxToHtmlConverter.Convert("notes.docx", new DocxToHtmlOptions
{
    OutputMode = HtmlOutputMode.BodyFragment
});
```

### DocumentTitle

Title element text. When null, the package title is used, then the file name, then `Document`.

```csharp
var result = DocxToHtmlConverter.Convert("notes.docx", new DocxToHtmlOptions
{
    DocumentTitle = "Field notes"
});
```

### CssMode and CssFileName

`CssMode.External` (default) links to `CssFileName`. `Embedded` puts the stylesheet in a `style` element. `None` leaves the HTML without a stylesheet reference; `Css` is still returned.

```csharp
var embedded = DocxToHtmlConverter.Convert("notes.docx", new DocxToHtmlOptions
{
    CssMode = CssMode.Embedded
});

var external = DocxToHtmlConverter.Convert("notes.docx", new DocxToHtmlOptions
{
    CssMode = CssMode.External,
    CssFileName = "notes.css"
});
```

### ClassPrefix

Prefix for generated classes. Mapped classes from `StyleMap` are not prefixed.

```csharp
var result = DocxToHtmlConverter.Convert("notes.docx", new DocxToHtmlOptions
{
    ClassPrefix = "dx-"
});
// <h6 class="dx-p-heading-7 dx-h7">Deep</h6>
```

### AdditionalCss

Text appended after the generated rules. It is not parsed.

```csharp
var result = DocxToHtmlConverter.Convert("notes.docx", new DocxToHtmlOptions
{
    AdditionalCss = "body { max-width: 40rem; margin: 2rem auto; }"
});
```

### IncludeBaseCss

When true (default), the stylesheet includes rules for body text, headings, paragraphs, and links. Tables, lists, notes, text boxes, and their padding are included only when the document uses them. Spacing that matches the document default is not copied onto every style. Set it false to emit only rules for styles and classes that appear.

```csharp
var result = DocxToHtmlConverter.Convert("notes.docx", new DocxToHtmlOptions
{
    IncludeBaseCss = false
});
```

### DirectFormatting

- `None` drops direct run and paragraph formatting. Character styles are kept.
- `SemanticOnly` emits `strong`, `em`, `u`, `s`, `sub`, and `sup`.
- `SemanticAndClasses` (default) adds classes for colour, highlight, size, and font. Direct spacing is a shared step such as `mt-12` or `ml-36`, not a separate class for every measurement.

Table shading, cell alignment, column widths, and image sizes are classes in every mode. They are not inline styles.

```csharp
var plain = DocxToHtmlConverter.Convert("notes.docx", new DocxToHtmlOptions
{
    DirectFormatting = DirectFormattingMode.SemanticOnly
});
```

### UseSemanticHeadings

When true (default), Heading styles, the Title style, and outline levels become `h1`–`h6`. When false, those paragraphs stay `p`. They are then omitted from `Headings`.

```csharp
var result = DocxToHtmlConverter.Convert("notes.docx", new DocxToHtmlOptions
{
    UseSemanticHeadings = false
});
```

### StyleMap

Overrides heading detection by Word style **name** (not style id), compared without case sensitivity. The class tokens are written as given.

```csharp
var result = DocxToHtmlConverter.Convert("notes.docx", new DocxToHtmlOptions
{
    StyleMap =
    {
        ["heading 1"] = new StyleMapping("p", "lead"),
        ["Quote"] = new StyleMapping("blockquote")
    }
});
```

A mapped tag above `h6`, such as `h7`, is written as `h6` and still receives the class `h7`.

### IncludeEmptyParagraphs

When false (default), paragraphs with no visible content are omitted. List items are kept. A bookmark id keeps an otherwise empty paragraph so the anchor remains.

```csharp
var result = DocxToHtmlConverter.Convert("notes.docx", new DocxToHtmlOptions
{
    IncludeEmptyParagraphs = true
});
```

### ImageMode and ImageFolder

`External` (default) uses `images/image-001.png`. `DataUri` uses a base64 `src`. `Skip` omits the picture and adds an `ImageSkipped` warning. Width and height are classes such as `img-w-320-h-160`, never attributes.

```csharp
var result = DocxToHtmlConverter.Convert("notes.docx", new DocxToHtmlOptions
{
    ImageMode = ImageMode.DataUri,
    ImageFolder = "media"
});
```

### IncludeHeadersAndFooters

When false (default), headers and footers are left out. When true, the default header and footer of the first section are included and their headings are collected.

```csharp
var result = DocxToHtmlConverter.Convert("notes.docx", new DocxToHtmlOptions
{
    IncludeHeadersAndFooters = true
});
```

### IncludeFootnotes

When true (default), referenced footnotes and endnotes are written at the end with a backlink. When false, the references and the notes are omitted.

```csharp
var result = DocxToHtmlConverter.Convert("notes.docx", new DocxToHtmlOptions
{
    IncludeFootnotes = false
});
```

### IncludeComments

When true, comment text is written at each comment reference in a `span` with class `comment`. Default is false.

```csharp
var result = DocxToHtmlConverter.Convert("notes.docx", new DocxToHtmlOptions
{
    IncludeComments = true
});
```

### TrackedChanges

`Accept` (default) keeps insertions and drops deletions. `Reject` does the opposite.

```csharp
var result = DocxToHtmlConverter.Convert("notes.docx", new DocxToHtmlOptions
{
    TrackedChanges = TrackedChangesMode.Reject
});
```

### OnWarning

Called as each recoverable problem is recorded. An exception from the callback is not caught.

```csharp
var result = DocxToHtmlConverter.Convert("notes.docx", new DocxToHtmlOptions
{
    OnWarning = warning => Console.Error.WriteLine($"{warning.Code}: {warning.Message}")
});
```

Warning codes include `UnknownElement`, `UnknownStyle`, `UnresolvedHyperlink`, `UnresolvedNumbering`, `MissingImage`, `ImageSkipped`, `UnsupportedSymbol`, `UnsupportedContent`, `MergedCellContent`, `OrphanedMerge`, `MissingNote`, `MissingComment`, `ThemeColor`, `UnclosedField`, `MissingHeaderFooter`, and `Highlight`.

## What the conversion keeps

Paragraph styles other than Normal become classes such as `p-quote`. Normal has no class. Runs that share the same visual formatting are merged. Hyperlinks and `HYPERLINK` fields become links; other fields keep their result and drop the instruction. Lists use `ul` or `ol`, with `start` when the numbering starts above 1. Tables keep `colspan`, `rowspan`, header rows, and nested tables. Pictures keep their alternative text. Text boxes become `div.textbox` and are not placed inside a paragraph.

## Command line

```bash
dotnet run --project samples/DocxToHtml.Cli -- notes.docx ./out --prefix dx- --headers
```

Flags: `--prefix`, `--title`, `--fragment`, `--embed-css`, `--headers`, `--comments`, `--reject-changes`, `--images external|data-uri|skip`.

## Tests

Fixtures live in `tests/DocxToHtml.Tests/Fixtures`. Snapshots live in `tests/DocxToHtml.Tests/Snapshots`. Every fixture is checked for three invariants: no element has a `style` attribute, every `span` has a class, and the visible document text matches the HTML text. Regenerate snapshots with:

```bash
UPDATE_SNAPSHOTS=1 dotnet test tests/DocxToHtml.Tests
```

## Limits

This is a semantic conversion, not a page layout. It does not paginate, draw the exact Word canvas, or edit the document. Theme colours without a resolved RGB value produce a warning. Levels above 9 are not a Word heading style; outline levels 0–8 map to headings 1–9.
