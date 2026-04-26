using System.Collections.Generic;
using System.Linq;
using System.Drawing;

namespace FastReport.Export.Custom;

public sealed class DocxLayoutSnapshot
{
    public float PageWidth { get; init; }

    public float PageHeight { get; init; }

    public float LeftMargin { get; init; }

    public float TopMargin { get; init; }

    public float RightMargin { get; init; }

    public float BottomMargin { get; init; }

    public IReadOnlyList<DocxBandSnapshot> Bands { get; init; } = new List<DocxBandSnapshot>();

    public IReadOnlyList<DocxTextSnapshot> Texts { get; init; } = new List<DocxTextSnapshot>();
}

public sealed class DocxBandSnapshot
{
    public string Name { get; init; } = string.Empty;

    public string BandType { get; init; } = string.Empty;

    public float Left { get; init; }

    public float Top { get; init; }

    public float Width { get; init; }

    public float Height { get; init; }

    public bool HasFill { get; init; }

    public bool HasBorder { get; init; }
}

public sealed class DocxTextSnapshot
{
    public string Name { get; init; } = string.Empty;

    public string Text { get; init; } = string.Empty;

    public string ParentBandName { get; init; } = string.Empty;

    public float Left { get; init; }

    public float Top { get; init; }

    public float Width { get; init; }

    public float Height { get; init; }

    public float PaddingLeft { get; init; }

    public float PaddingTop { get; init; }

    public float PaddingRight { get; init; }

    public float PaddingBottom { get; init; }

    public HorzAlign HorzAlign { get; init; }

    public string FontName { get; init; } = string.Empty;

    public float FontSizeInPoints { get; init; }

    public bool Bold { get; init; }

    public bool Italic { get; init; }

    public bool Underline { get; init; }

    public string TextColor { get; init; } = string.Empty;

    public string FillColor { get; init; } = string.Empty;

    public string BorderColor { get; init; } = string.Empty;

    public bool HasFill { get; init; }

    public bool HasBorder { get; init; }
}

public static class DocxLayoutSnapshotBuilder
{
    public static DocxLayoutSnapshot Build(ReportPage page)
    {
        List<DocxBandSnapshot> bands = page.AllObjects
            .OfType<BandBase>()
            .Select(band => new DocxBandSnapshot
            {
                Name = band.Name,
                BandType = band.GetType().Name,
                Left = band.AbsLeft,
                Top = band.AbsTop,
                Width = band.Width,
                Height = band.Height,
                HasFill = band.Fill is not SolidFill solidFill || !solidFill.IsTransparent,
                HasBorder = band.Border?.Lines != BorderLines.None
            })
            .OrderBy(band => band.Top)
            .ThenBy(band => band.Left)
            .ToList();

        List<DocxTextSnapshot> texts = page.AllObjects
            .OfType<TextObjectBase>()
            .Select(text => new DocxTextSnapshot
            {
                Name = text.Name,
                Text = text.Text ?? string.Empty,
                ParentBandName = text.Parent is BandBase band ? band.Name : string.Empty,
                Left = text.AbsLeft,
                Top = text.AbsTop,
                Width = text.Width,
                Height = text.Height,
                PaddingLeft = text.Padding.Left,
                PaddingTop = text.Padding.Top,
                PaddingRight = text.Padding.Right,
                PaddingBottom = text.Padding.Bottom,
                HorzAlign = text is TextObject textObject ? textObject.HorzAlign : HorzAlign.Left,
                FontName = text is TextObject styledText ? styledText.Font.Name : string.Empty,
                FontSizeInPoints = text is TextObject sizedText ? sizedText.Font.SizeInPoints : 0,
                Bold = text is TextObject boldText && boldText.Font.Bold,
                Italic = text is TextObject italicText && italicText.Font.Italic,
                Underline = text is TextObject underlinedText && underlinedText.Underlines,
                TextColor = text is TextObject coloredText ? ToHexColor(GetColorFromFill(coloredText.TextFill)) : string.Empty,
                FillColor = ToHexColor(GetColorFromFill(text.Fill)),
                BorderColor = ToHexColor(text.Border?.Color ?? Color.Transparent),
                HasFill = text.Fill is not SolidFill solidFill || !solidFill.IsTransparent,
                HasBorder = text.Border?.Lines != BorderLines.None
            })
            .OrderBy(text => text.Top)
            .ThenBy(text => text.Left)
            .ToList();

        return new DocxLayoutSnapshot
        {
            PageWidth = page.WidthInPixels,
            PageHeight = page.HeightInPixels,
            LeftMargin = page.LeftMargin,
            TopMargin = page.TopMargin,
            RightMargin = page.RightMargin,
            BottomMargin = page.BottomMargin,
            Bands = bands,
            Texts = texts
        };
    }

    private static string ToHexColor(Color color)
    {
        return color.A == 0 ? string.Empty : $"{color.R:X2}{color.G:X2}{color.B:X2}";
    }

    private static Color GetColorFromFill(FillBase? fill)
    {
        return fill switch
        {
            SolidFill solidFill => solidFill.Color,
            LinearGradientFill gradientFill => gradientFill.StartColor,
            PathGradientFill pathGradientFill => pathGradientFill.CenterColor,
            HatchFill hatchFill => hatchFill.ForeColor,
            GlassFill glassFill => glassFill.Color,
            _ => Color.Transparent
        };
    }
}
