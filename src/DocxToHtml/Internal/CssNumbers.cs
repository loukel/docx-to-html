namespace DocxToHtml.Internal;

internal static class CssNumbers
{
    public static string Points(double points) =>
        points.ToString("0.##", CultureInfo.InvariantCulture);

    public static string Decimal(double value) =>
        value.ToString("0.###", CultureInfo.InvariantCulture);

    public static string PointToken(double points) =>
        Points(points).Replace('.', '-');
}
