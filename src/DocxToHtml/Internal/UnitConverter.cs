namespace DocxToHtml.Internal;

internal static class UnitConverter
{
    public const int TwipsPerPoint = 20;
    public const int EmuPerInch = 914400;
    public const int CssPixelsPerInch = 96;
    public const int LineUnitPerSingleSpacing = 240;

    public static int EmuToPixels(long englishMetricUnits)
    {
        return (int)Math.Round(
            englishMetricUnits / (double)EmuPerInch * CssPixelsPerInch,
            MidpointRounding.AwayFromZero);
    }

    public static int PointsToPixels(double points)
    {
        return (int)Math.Round(points * CssPixelsPerInch / 72d, MidpointRounding.AwayFromZero);
    }

    public static double TwipsToPoints(int twips) => twips / (double)TwipsPerPoint;

    public static double HalfPointsToPoints(int halfPoints) => halfPoints / 2d;
}
