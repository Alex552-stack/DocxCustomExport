using System;

namespace FastReport.Export.Custom;

internal static class DocxUnitConverter
{
    private const double TwipsPerInch = 1440d;
    private const double EmusPerInch = 914400d;
    private const double MillimetersPerInch = 25.4d;
    private const double PixelsPerInch = 96d;

    public static UInt32 MillimetersToTwips(float millimeters)
    {
        if (millimeters < 0)
            throw new ArgumentOutOfRangeException(nameof(millimeters));

        return (UInt32)Math.Round(millimeters * TwipsPerInch / MillimetersPerInch, MidpointRounding.AwayFromZero);
    }

    public static UInt32 MillimetersToEmus(float millimeters)
    {
        if (millimeters < 0)
            throw new ArgumentOutOfRangeException(nameof(millimeters));

        return (UInt32)Math.Round(millimeters * EmusPerInch / MillimetersPerInch, MidpointRounding.AwayFromZero);
    }

    public static Int32 PixelsToTwips(float pixels)
    {
        return (Int32)Math.Round(pixels * TwipsPerInch / PixelsPerInch, MidpointRounding.AwayFromZero);
    }

    public static Int64 PixelsToEmus(float pixels)
    {
        return (Int64)Math.Round(pixels * EmusPerInch / PixelsPerInch, MidpointRounding.AwayFromZero);
    }

    public static float MillimetersToPixels(float millimeters)
    {
        if (millimeters < 0)
            throw new ArgumentOutOfRangeException(nameof(millimeters));

        return (float)(millimeters * PixelsPerInch / MillimetersPerInch);
    }

    public static Int32 PixelsToPoints(float pixels)
    {
        return (Int32)Math.Round(pixels * 72d / PixelsPerInch, MidpointRounding.AwayFromZero);
    }
}
