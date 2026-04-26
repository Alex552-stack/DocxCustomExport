using FastReport.Table;
using FastReport.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;

namespace FastReport.Export.Custom;

internal sealed class DocxMatrixPage
{
    public float PageWidth { get; init; }

    public float PageHeight { get; init; }

    public float LeftMargin { get; init; }

    public float TopMargin { get; init; }

    public float RightMargin { get; init; }

    public float BottomMargin { get; init; }

    public bool Landscape { get; init; }

    public IReadOnlyList<float> XEdges { get; init; } = Array.Empty<float>();

    public IReadOnlyList<float> YEdges { get; init; } = Array.Empty<float>();

    public IReadOnlyList<DocxMatrixObject> Objects { get; init; } = Array.Empty<DocxMatrixObject>();

    public IReadOnlyList<DocxMatrixFragment> Fragments { get; init; } = Array.Empty<DocxMatrixFragment>();

    public int ColumnCount => Math.Max(0, XEdges.Count - 1);

    public int RowCount => Math.Max(0, YEdges.Count - 1);
}

internal sealed class DocxMatrixRegions
{
    public DocxMatrixPage Body { get; init; } = new();

    public DocxMatrixPage? Header { get; init; }

    public DocxMatrixPage? Footer { get; init; }
}

internal enum DocxPageRegion
{
    Body,
    Header,
    Footer
}

internal sealed class DocxMatrixObject
{
    public string Name { get; init; } = string.Empty;

    public string Kind { get; init; } = string.Empty;

    public string ParentBandName { get; init; } = string.Empty;

    public float Left { get; init; }

    public float Top { get; init; }

    public float Width { get; init; }

    public float Height { get; init; }

    public bool IsText { get; init; }

    public string Text { get; init; } = string.Empty;

    public string Bookmark { get; init; } = string.Empty;

    public HyperlinkKind? HyperlinkKind { get; init; }

    public string Hyperlink { get; init; } = string.Empty;

    public byte[]? ImageBytes { get; init; }

    public string ImageContentType { get; init; } = string.Empty;

    public float ImageWidth { get; init; }

    public float ImageHeight { get; init; }

    public int StyleIndex { get; init; }

    public DocxMatrixStyle Style { get; init; } = DocxMatrixStyle.Empty;

    public int Column { get; init; }

    public int Row { get; init; }

    public int ColumnSpan { get; init; }

    public int RowSpan { get; init; }

    public bool HasImage => ImageBytes != null && ImageBytes.Length > 0;

    public bool ShouldOccupyCells => IsText || HasImage || Style.HasFill || Style.HasBorder;
}

internal sealed class DocxMatrixFragment
{
    public int ObjectIndex { get; init; }

    public int Column { get; init; }

    public int Row { get; init; }

    public int ColumnSpan { get; init; }

    public int RowSpan { get; init; }
}

internal sealed class DocxMatrixStyle : IEquatable<DocxMatrixStyle>
{
    public static DocxMatrixStyle Empty { get; } = new();

    public string FillColor { get; init; } = string.Empty;

    public string BorderColor { get; init; } = string.Empty;

    public bool HasExplicitFill { get; init; }

    public bool HasExplicitBorder { get; init; }

    public string TextColor { get; init; } = string.Empty;

    public string FontName { get; init; } = string.Empty;

    public float FontSizeInPoints { get; init; }

    public bool Bold { get; init; }

    public bool Italic { get; init; }

    public bool Underline { get; init; }

    public HorzAlign HorzAlign { get; init; } = HorzAlign.Left;

    public VertAlign VertAlign { get; init; } = VertAlign.Top;

    public bool RightToLeft { get; init; }

    public bool WordWrap { get; init; } = true;

    public float PaddingLeft { get; init; }

    public float PaddingTop { get; init; }

    public float PaddingRight { get; init; }

    public float PaddingBottom { get; init; }

    public float ParagraphOffset { get; init; }

    public float LineHeight { get; init; }

    public bool HasFill => HasExplicitFill && !string.IsNullOrEmpty(FillColor);

    public bool HasBorder => HasExplicitBorder && !string.IsNullOrEmpty(BorderColor);

    public DocxMatrixStyle WithInheritedFrame(DocxMatrixStyle background)
    {
        if (background == null)
            return this;

        return new DocxMatrixStyle
        {
            FillColor = HasFill ? FillColor : background.FillColor,
            BorderColor = HasBorder ? BorderColor : background.BorderColor,
            HasExplicitFill = HasFill || background.HasFill,
            HasExplicitBorder = HasBorder || background.HasBorder,
            TextColor = TextColor,
            FontName = FontName,
            FontSizeInPoints = FontSizeInPoints,
            Bold = Bold,
            Italic = Italic,
            Underline = Underline,
            HorzAlign = HorzAlign,
            VertAlign = VertAlign,
            RightToLeft = RightToLeft,
            WordWrap = WordWrap,
            PaddingLeft = PaddingLeft,
            PaddingTop = PaddingTop,
            PaddingRight = PaddingRight,
            PaddingBottom = PaddingBottom,
            ParagraphOffset = ParagraphOffset,
            LineHeight = LineHeight
        };
    }

    public bool Equals(DocxMatrixStyle? other)
    {
        if (other == null)
            return false;

        return FillColor == other.FillColor &&
            BorderColor == other.BorderColor &&
            HasExplicitFill == other.HasExplicitFill &&
            HasExplicitBorder == other.HasExplicitBorder &&
            TextColor == other.TextColor &&
            FontName == other.FontName &&
            FontSizeInPoints.Equals(other.FontSizeInPoints) &&
            Bold == other.Bold &&
            Italic == other.Italic &&
            Underline == other.Underline &&
            HorzAlign == other.HorzAlign &&
            VertAlign == other.VertAlign &&
            RightToLeft == other.RightToLeft &&
            WordWrap == other.WordWrap &&
            PaddingLeft.Equals(other.PaddingLeft) &&
            PaddingTop.Equals(other.PaddingTop) &&
            PaddingRight.Equals(other.PaddingRight) &&
            PaddingBottom.Equals(other.PaddingBottom) &&
            ParagraphOffset.Equals(other.ParagraphOffset) &&
            LineHeight.Equals(other.LineHeight);
    }

    public override bool Equals(object? obj)
    {
        return Equals(obj as DocxMatrixStyle);
    }

    public override int GetHashCode()
    {
        HashCode hash = new();
        hash.Add(FillColor);
        hash.Add(BorderColor);
        hash.Add(HasExplicitFill);
        hash.Add(HasExplicitBorder);
        hash.Add(TextColor);
        hash.Add(FontName);
        hash.Add(FontSizeInPoints);
        hash.Add(Bold);
        hash.Add(Italic);
        hash.Add(Underline);
        hash.Add(HorzAlign);
        hash.Add(VertAlign);
        hash.Add(RightToLeft);
        hash.Add(WordWrap);
        hash.Add(PaddingLeft);
        hash.Add(PaddingTop);
        hash.Add(PaddingRight);
        hash.Add(PaddingBottom);
        hash.Add(ParagraphOffset);
        hash.Add(LineHeight);
        return hash.ToHashCode();
    }
}

internal static class DocxMatrixBuilder
{
    private const float DefaultTolerance = 0.5f;
    private const float PaginationSafetyTrimPixels = 0.25f;
    private const float PaginationSafetyThresholdPixels = 1f;

    public static DocxMatrixPage Build(ReportPage page, float tolerance = DefaultTolerance)
    {
        List<DocxMatrixStyle> styles = new();
        List<CollectedObject> collectedObjects = CollectObjects(page, styles);
        return BuildRegionPage(page, collectedObjects, DocxPageRegion.Body, tolerance) ?? CreateEmptyPage(page);
    }

    public static DocxMatrixRegions BuildRegions(ReportPage page, float tolerance = DefaultTolerance)
    {
        List<DocxMatrixStyle> styles = new();
        List<CollectedObject> collectedObjects = CollectObjects(page, styles);

        return new DocxMatrixRegions
        {
            Body = BuildRegionPage(page, collectedObjects, DocxPageRegion.Body, tolerance) ?? CreateEmptyPage(page),
            Header = BuildRegionPage(page, collectedObjects, DocxPageRegion.Header, tolerance),
            Footer = BuildRegionPage(page, collectedObjects, DocxPageRegion.Footer, tolerance)
        };
    }

    private static float GetContentExtent(IEnumerable<float> edges, float fallback)
    {
        float extent = edges.DefaultIfEmpty(0).Max();
        return extent > 0 ? extent : fallback;
    }

    private static float ReduceTrailingExtent(float contentExtent, ReportPage page)
    {
        if (contentExtent <= 0 || page.HeightInPixels <= 0)
            return contentExtent;

        if (Math.Abs(contentExtent - page.HeightInPixels) < 0.5f)
            contentExtent = Math.Max(0, contentExtent - 1f);

        float printableHeight = GetPrintableHeightPixels(page);
        if (printableHeight > 0 &&
            contentExtent >= printableHeight - PaginationSafetyThresholdPixels &&
            contentExtent <= printableHeight + PaginationSafetyThresholdPixels)
        {
            return Math.Max(0, printableHeight - PaginationSafetyTrimPixels);
        }

        return contentExtent;
    }

    private static float GetPrintableHeightPixels(ReportPage page)
    {
        float topMargin = DocxUnitConverter.MillimetersToPixels(page.TopMargin);
        float bottomMargin = DocxUnitConverter.MillimetersToPixels(page.BottomMargin);
        return Math.Max(0, page.HeightInPixels - topMargin - bottomMargin);
    }

    private static List<CollectedObject> CollectObjects(ReportPage page, List<DocxMatrixStyle> styles)
    {
        List<CollectedObject> result = new();
        List<BandBase> bands = page.AllObjects
            .OfType<BandBase>()
            .OrderBy(x => x.AbsTop)
            .ThenBy(x => x.AbsLeft)
            .ToList();

        foreach (BandBase band in bands)
        {
            if (band.Fill is not SolidFill solidFill || !solidFill.IsTransparent || band.Border?.Lines != BorderLines.None)
                result.Add(CreateBandObject(page, band, styles));

            foreach (Base item in band.ForEachAllConvectedObjects(page))
            {
                if (item is not ReportComponentBase component || !component.Exportable)
                    continue;

                if (component is CellularTextObject cellularText)
                    component = (ReportComponentBase)cellularText.GetTable();

                if (component is TableCell)
                    continue;

                if (component is TableBase table)
                {
                    result.AddRange(CreateTableObjects(page, table, styles, band));
                    continue;
                }

                if (component is TextObjectBase text)
                {
                    CollectedObject? textObject = CreateTextObjectOrImage(page, text, styles, text.AbsLeft, text.AbsTop, band);
                    if (textObject != null)
                        result.Add(textObject);
                    continue;
                }

                if (component is PictureObject picture)
                {
                    CollectedObject? imageObject = CreateImageObject(page, picture, styles, band);
                    if (imageObject != null)
                        result.Add(imageObject);
                    continue;
                }

                if (component is BandBase nestedBand)
                {
                    if (nestedBand.Fill is not SolidFill nestedFill || !nestedFill.IsTransparent || nestedBand.Border?.Lines != BorderLines.None)
                        result.Add(CreateBandObject(page, nestedBand, styles));
                    continue;
                }

                CollectedObject? renderedObject = CreateRenderedComponentObject(page, component, styles, band);
                if (renderedObject != null)
                    result.Add(renderedObject);
            }
        }

        return result
            .Where(x => x.Width > 0 && x.Height > 0)
            .ToList();
    }

    private static DocxMatrixPage? BuildRegionPage(
        ReportPage page,
        List<CollectedObject> collectedObjects,
        DocxPageRegion region,
        float tolerance)
    {
        List<CollectedObject> regionObjects = NormalizeRegionObjects(collectedObjects, region);
        if (regionObjects.Count == 0)
            return null;

        float pageHeightFallback = region == DocxPageRegion.Body ? page.HeightInPixels : regionObjects.Max(x => x.Top + x.Height);
        float contentWidth = GetContentExtent(regionObjects.Select(x => x.Left + x.Width), page.WidthInPixels);
        float contentHeight = GetContentExtent(regionObjects.Select(x => x.Top + x.Height), pageHeightFallback);
        if (region == DocxPageRegion.Body)
            contentHeight = ReduceTrailingExtent(contentHeight, page);

        List<float> xEdges = BuildEdges(contentWidth, tolerance, regionObjects.SelectMany(ToHorizontalEdges));
        List<float> yEdges = BuildEdges(contentHeight, tolerance, regionObjects.SelectMany(ToVerticalEdges));

        List<DocxMatrixObject> objects = regionObjects
            .Select(item => CreateMatrixObject(item, xEdges, yEdges, tolerance))
            .ToList();

        int[] occupancy = Rasterize(objects, xEdges.Count - 1, yEdges.Count - 1);
        List<DocxMatrixFragment> fragments = AnalyzeFragments(occupancy, xEdges.Count - 1, yEdges.Count - 1);

        return new DocxMatrixPage
        {
            PageWidth = contentWidth,
            PageHeight = contentHeight,
            LeftMargin = page.LeftMargin,
            TopMargin = page.TopMargin,
            RightMargin = page.RightMargin,
            BottomMargin = page.BottomMargin,
            Landscape = page.Landscape,
            XEdges = xEdges,
            YEdges = yEdges,
            Objects = objects,
            Fragments = fragments
        };
    }

    private static List<CollectedObject> NormalizeRegionObjects(List<CollectedObject> collectedObjects, DocxPageRegion region)
    {
        List<CollectedObject> regionObjects = collectedObjects
            .Where(x => x.Region == region)
            .ToList();

        if (regionObjects.Count == 0)
            return regionObjects;

        float originTop = region == DocxPageRegion.Body ? 0 : regionObjects.Min(x => x.Top);
        if (originTop <= 0)
            return regionObjects;

        return regionObjects
            .Select(item => CloneCollectedObject(item, item.Left, item.Top - originTop))
            .ToList();
    }

    private static DocxMatrixPage CreateEmptyPage(ReportPage page)
    {
        return new DocxMatrixPage
        {
            PageWidth = 0,
            PageHeight = 0,
            LeftMargin = page.LeftMargin,
            TopMargin = page.TopMargin,
            RightMargin = page.RightMargin,
            BottomMargin = page.BottomMargin,
            Landscape = page.Landscape,
            XEdges = new[] { 0f },
            YEdges = new[] { 0f }
        };
    }

    private static IEnumerable<CollectedObject> CreateTableObjects(ReportPage page, TableBase table, List<DocxMatrixStyle> styles, BandBase parentBand)
    {
        float top = 0;
        for (int rowIndex = 0; rowIndex < table.RowCount; rowIndex++)
        {
            float left = 0;
            for (int columnIndex = 0; columnIndex < table.ColumnCount; columnIndex++)
            {
                TableCell cell = table[columnIndex, rowIndex];
                if (!table.IsInsideSpan(cell))
                {
                    yield return CreateTextObject(page, cell, styles, table.AbsLeft + left, table.AbsTop + top, parentBand);
                }

                left += table.Columns[columnIndex].Width;
            }

            top += table.Rows[rowIndex].Height;
        }
    }

    private static CollectedObject CreateBandObject(ReportPage page, BandBase band, List<DocxMatrixStyle> styles)
    {
        DocxMatrixStyle style = CreateFrameStyle(band);
        return new CollectedObject
        {
            Name = band.Name,
            Kind = band.GetType().Name,
            ParentBandName = string.Empty,
            Left = band.AbsLeft,
            Top = band.AbsTop,
            Width = band.Width,
            Height = band.Height,
            Region = GetRegion(page, band, band.AbsTop, band.Height),
            Style = style,
            StyleIndex = GetStyleIndex(styles, style),
            IsText = false
        };
    }

    private static CollectedObject? CreateTextObjectOrImage(ReportPage page, TextObjectBase text, List<DocxMatrixStyle> styles, float left, float top, BandBase? parentBand = null)
    {
        if (ShouldRenderTextAsImage(text))
            return CreateRenderedComponentObject(page, text, styles, parentBand, left, top);

        return CreateTextObject(page, text, styles, left, top, parentBand);
    }

    private static CollectedObject CreateTextObject(ReportPage page, TextObjectBase text, List<DocxMatrixStyle> styles, float left, float top, BandBase? parentBand = null)
    {
        DocxMatrixStyle style = CreateTextStyle(text);
        BandBase? ownerBand = parentBand ?? text.Parent as BandBase;
        return new CollectedObject
        {
            Name = text.Name,
            Kind = nameof(TextObjectBase),
            ParentBandName = ownerBand?.Name ?? string.Empty,
            Left = left,
            Top = top,
            Width = text.Width,
            Height = text.Height,
            IsText = true,
            Text = text.Text ?? string.Empty,
            Bookmark = text.Bookmark ?? string.Empty,
            HyperlinkKind = string.IsNullOrWhiteSpace(text.Hyperlink.Value) ? null : text.Hyperlink.Kind,
            Hyperlink = text.Hyperlink.Value ?? string.Empty,
            Region = GetRegion(page, ownerBand, top, text.Height),
            Style = style,
            StyleIndex = GetStyleIndex(styles, style)
        };
    }

    private static CollectedObject? CreateImageObject(ReportPage page, PictureObject picture, List<DocxMatrixStyle> styles, BandBase? parentBand)
    {
        byte[]? imageBytes = TryGetImageBytes(picture);
        if (imageBytes == null || imageBytes.Length == 0)
            return null;

        DocxMatrixStyle style = CreateFrameStyle(picture);
        return new CollectedObject
        {
            Name = picture.Name,
            Kind = nameof(PictureObject),
            ParentBandName = parentBand?.Name ?? string.Empty,
            Left = picture.AbsLeft,
            Top = picture.AbsTop,
            Width = picture.Width,
            Height = picture.Height,
            Style = style,
            StyleIndex = GetStyleIndex(styles, style),
            Region = GetRegion(page, parentBand, picture.AbsTop, picture.Height),
            ImageBytes = imageBytes,
            ImageContentType = "image/png",
            ImageWidth = picture.Width,
            ImageHeight = picture.Height
        };
    }

    private static CollectedObject? CreateRenderedComponentObject(
        ReportPage page,
        ReportComponentBase component,
        List<DocxMatrixStyle> styles,
        BandBase? parentBand,
        float? leftOverride = null,
        float? topOverride = null)
    {
        byte[]? imageBytes = RenderComponentToPng(component);
        if (imageBytes == null || imageBytes.Length == 0)
            return null;

        DocxMatrixStyle style = CreateFrameStyle(component);
        return new CollectedObject
        {
            Name = component.Name,
            Kind = component.GetType().Name,
            ParentBandName = parentBand?.Name ?? string.Empty,
            Left = leftOverride ?? component.AbsLeft,
            Top = topOverride ?? component.AbsTop,
            Width = component.Width,
            Height = component.Height,
            Style = style,
            StyleIndex = GetStyleIndex(styles, style),
            Region = GetRegion(page, parentBand, topOverride ?? component.AbsTop, component.Height),
            ImageBytes = imageBytes,
            ImageContentType = "image/png",
            ImageWidth = component.Width,
            ImageHeight = component.Height
        };
    }

    private static DocxPageRegion GetRegion(ReportPage page, BandBase? band, float top, float height)
    {
        if (band is PageHeaderBand)
            return DocxPageRegion.Header;

        if (band is PageFooterBand)
            return DocxPageRegion.Footer;

        return DocxPageRegion.Body;
    }

    private static byte[]? TryGetImageBytes(PictureObject picture)
    {
        return RenderComponentToPng(picture) ?? SaveImageToPng(picture.Image);
    }

    private static bool ShouldRenderTextAsImage(TextObjectBase text)
    {
        return text is TextObject textObject && textObject.Angle != 0;
    }

    private static byte[]? RenderComponentToPng(ReportComponentBase component)
    {
        int width = Math.Max(1, (int)Math.Ceiling(component.Width));
        int height = Math.Max(1, (int)Math.Ceiling(component.Height));
        using Bitmap bitmap = new(width, height);
        using Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.Transparent);
        graphics.TranslateTransform(-component.AbsLeft, -component.AbsTop);

        GraphicCache? ownedCache = null;
        GraphicCache cache = component.Report?.GraphicCache ?? (ownedCache = new GraphicCache());
        try
        {
            component.Draw(new FRPaintEventArgs(graphics, 1f, 1f, cache));
        }
        catch
        {
            return null;
        }
        finally
        {
            ownedCache?.Dispose();
        }

        return SaveImageToPng(bitmap);
    }

    private static byte[]? SaveImageToPng(System.Drawing.Image? image)
    {
        if (image == null)
            return null;

        using MemoryStream stream = new();
        if (image is Bitmap)
        {
            image.Save(stream, ImageFormat.Png);
            return stream.ToArray();
        }

        using Bitmap bitmap = new(image.Width, image.Height);
        using Graphics graphics = Graphics.FromImage(bitmap);
        graphics.DrawImage(image, 0, 0, image.Width, image.Height);
        bitmap.Save(stream, ImageFormat.Png);
        return stream.ToArray();
    }

    private static DocxMatrixStyle CreateFrameStyle(ReportComponentBase component)
    {
        return new DocxMatrixStyle
        {
            FillColor = ToHexColor(GetColorFromFill(component.Fill)),
            BorderColor = ToHexColor(component.Border?.Color ?? Color.Transparent),
            HasExplicitFill = component.Fill is not SolidFill solidFill || !solidFill.IsTransparent,
            HasExplicitBorder = component.Border?.Lines != BorderLines.None
        };
    }

    private static DocxMatrixStyle CreateTextStyle(TextObjectBase text)
    {
        if (text is not TextObject textObject)
        {
            return new DocxMatrixStyle
            {
                FillColor = ToHexColor(GetColorFromFill(text.Fill)),
                BorderColor = ToHexColor(text.Border?.Color ?? Color.Transparent),
                HasExplicitFill = text.Fill is not SolidFill solidFill || !solidFill.IsTransparent,
                HasExplicitBorder = text.Border?.Lines != BorderLines.None,
                PaddingLeft = text.Padding.Left,
                PaddingTop = text.Padding.Top,
                PaddingRight = text.Padding.Right,
                PaddingBottom = text.Padding.Bottom
            };
        }

        return new DocxMatrixStyle
        {
            FillColor = ToHexColor(GetColorFromFill(text.Fill)),
            BorderColor = ToHexColor(text.Border?.Color ?? Color.Transparent),
            HasExplicitFill = text.Fill is not SolidFill textFill || !textFill.IsTransparent,
            HasExplicitBorder = text.Border?.Lines != BorderLines.None,
            TextColor = ToHexColor(GetColorFromFill(textObject.TextFill)),
            FontName = textObject.Font.Name,
            FontSizeInPoints = textObject.Font.SizeInPoints,
            Bold = textObject.Font.Bold,
            Italic = textObject.Font.Italic,
            Underline = textObject.Font.Underline || textObject.Underlines,
            HorzAlign = textObject.HorzAlign,
            VertAlign = textObject.VertAlign,
            RightToLeft = textObject.RightToLeft,
            WordWrap = textObject.WordWrap,
            PaddingLeft = text.Padding.Left,
            PaddingTop = text.Padding.Top,
            PaddingRight = text.Padding.Right,
            PaddingBottom = text.Padding.Bottom,
            ParagraphOffset = textObject.ParagraphOffset,
            LineHeight = textObject.LineHeight
        };
    }

    private static int GetStyleIndex(List<DocxMatrixStyle> styles, DocxMatrixStyle style)
    {
        for (int index = styles.Count - 1; index >= 0; index--)
        {
            if (styles[index].Equals(style))
                return index;
        }

        styles.Add(style);
        return styles.Count - 1;
    }

    private static DocxMatrixObject CreateMatrixObject(
        CollectedObject item,
        IReadOnlyList<float> xEdges,
        IReadOnlyList<float> yEdges,
        float tolerance)
    {
        return new DocxMatrixObject
        {
            Name = item.Name,
            Kind = item.Kind,
            ParentBandName = item.ParentBandName,
            Left = item.Left,
            Top = item.Top,
            Width = item.Width,
            Height = item.Height,
            IsText = item.IsText,
            Text = item.Text,
            Bookmark = item.Bookmark,
            HyperlinkKind = item.HyperlinkKind,
            Hyperlink = item.Hyperlink,
            ImageBytes = item.ImageBytes,
            ImageContentType = item.ImageContentType,
            ImageWidth = item.ImageWidth,
            ImageHeight = item.ImageHeight,
            Style = item.Style,
            StyleIndex = item.StyleIndex,
            Column = FindEdgeIndex(xEdges, item.Left, tolerance),
            Row = FindEdgeIndex(yEdges, item.Top, tolerance),
            ColumnSpan = FindSpan(xEdges, item.Left, item.Width, tolerance),
            RowSpan = FindSpan(yEdges, item.Top, item.Height, tolerance)
        };
    }

    private static int[] Rasterize(List<DocxMatrixObject> objects, int columnCount, int rowCount)
    {
        int[] occupancy = Enumerable.Repeat(-1, Math.Max(0, columnCount * rowCount)).ToArray();

        for (int objectIndex = 0; objectIndex < objects.Count; objectIndex++)
        {
            DocxMatrixObject current = objects[objectIndex];
            if (!current.ShouldOccupyCells)
                continue;

            if (current.IsText)
            {
                int originIndex = current.Row * columnCount + current.Column;
                if (originIndex >= 0 && originIndex < occupancy.Length)
                {
                    int underlyingIndex = occupancy[originIndex];
                    if (underlyingIndex >= 0)
                    {
                        DocxMatrixStyle inheritedStyle = current.Style.WithInheritedFrame(objects[underlyingIndex].Style);
                        if (!ReferenceEquals(inheritedStyle, current.Style))
                        {
                            objects[objectIndex] = CloneObject(current, inheritedStyle);
                            current = objects[objectIndex];
                        }
                    }
                }
            }

            for (int row = current.Row; row < Math.Min(rowCount, current.Row + current.RowSpan); row++)
            {
                for (int column = current.Column; column < Math.Min(columnCount, current.Column + current.ColumnSpan); column++)
                {
                    occupancy[row * columnCount + column] = objectIndex;
                }
            }
        }

        return occupancy;
    }

    private static DocxMatrixObject CloneObject(DocxMatrixObject source, DocxMatrixStyle style)
    {
        return new DocxMatrixObject
        {
            Name = source.Name,
            Kind = source.Kind,
            ParentBandName = source.ParentBandName,
            Left = source.Left,
            Top = source.Top,
            Width = source.Width,
            Height = source.Height,
            IsText = source.IsText,
            Text = source.Text,
            Bookmark = source.Bookmark,
            HyperlinkKind = source.HyperlinkKind,
            Hyperlink = source.Hyperlink,
            ImageBytes = source.ImageBytes,
            ImageContentType = source.ImageContentType,
            ImageWidth = source.ImageWidth,
            ImageHeight = source.ImageHeight,
            StyleIndex = source.StyleIndex,
            Style = style,
            Column = source.Column,
            Row = source.Row,
            ColumnSpan = source.ColumnSpan,
            RowSpan = source.RowSpan
        };
    }

    private static List<DocxMatrixFragment> AnalyzeFragments(int[] occupancy, int columnCount, int rowCount)
    {
        bool[] visited = new bool[occupancy.Length];
        List<DocxMatrixFragment> fragments = new();

        for (int row = 0; row < rowCount; row++)
        {
            for (int column = 0; column < columnCount; column++)
            {
                int index = row * columnCount + column;
                int objectIndex = occupancy[index];
                if (objectIndex < 0 || visited[index])
                    continue;

                int width = 1;
                while (column + width < columnCount && occupancy[row * columnCount + column + width] == objectIndex && !visited[row * columnCount + column + width])
                    width++;

                int height = 1;
                bool expand = true;
                while (expand && row + height < rowCount)
                {
                    for (int x = 0; x < width; x++)
                    {
                        int probe = (row + height) * columnCount + column + x;
                        if (occupancy[probe] != objectIndex || visited[probe])
                        {
                            expand = false;
                            break;
                        }
                    }

                    if (expand)
                        height++;
                }

                for (int y = 0; y < height; y++)
                {
                    for (int x = 0; x < width; x++)
                    {
                        visited[(row + y) * columnCount + column + x] = true;
                    }
                }

                fragments.Add(new DocxMatrixFragment
                {
                    ObjectIndex = objectIndex,
                    Column = column,
                    Row = row,
                    ColumnSpan = width,
                    RowSpan = height
                });
            }
        }

        return fragments;
    }

    private static IEnumerable<float> ToHorizontalEdges(CollectedObject item)
    {
        yield return item.Left;
        yield return item.Left + item.Width;
    }

    private static IEnumerable<float> ToVerticalEdges(CollectedObject item)
    {
        yield return item.Top;
        yield return item.Top + item.Height;
    }

    private static List<float> BuildEdges(float pageExtent, float tolerance, IEnumerable<float> edges)
    {
        List<float> ordered = edges
            .Append(0)
            .Append(pageExtent)
            .Where(x => x >= 0 && x <= pageExtent)
            .OrderBy(x => x)
            .ToList();

        List<float> merged = new();
        foreach (float edge in ordered)
        {
            if (merged.Count == 0 || Math.Abs(edge - merged[^1]) > tolerance)
                merged.Add(edge);
        }

        if (merged.Count == 0 || merged[0] > tolerance)
            merged.Insert(0, 0);

        if (Math.Abs(merged[^1] - pageExtent) > tolerance)
            merged.Add(pageExtent);

        return merged;
    }

    private static int FindEdgeIndex(IReadOnlyList<float> edges, float value, float tolerance)
    {
        for (int index = 0; index < edges.Count; index++)
        {
            if (Math.Abs(edges[index] - value) <= tolerance)
                return index;
        }

        throw new InvalidOperationException($"Edge {value} was not found in the matrix.");
    }

    private static int FindSpan(IReadOnlyList<float> edges, float start, float extent, float tolerance)
    {
        int startIndex = FindEdgeIndex(edges, start, tolerance);
        int endIndex = FindEdgeIndex(edges, start + extent, tolerance);
        return Math.Max(1, endIndex - startIndex);
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

    private static string ToHexColor(Color color)
    {
        return color.A == 0 ? string.Empty : $"{color.R:X2}{color.G:X2}{color.B:X2}";
    }

    private static CollectedObject CloneCollectedObject(CollectedObject source, float left, float top)
    {
        return new CollectedObject
        {
            Name = source.Name,
            Kind = source.Kind,
            ParentBandName = source.ParentBandName,
            Left = left,
            Top = top,
            Width = source.Width,
            Height = source.Height,
            IsText = source.IsText,
            Text = source.Text,
            Bookmark = source.Bookmark,
            HyperlinkKind = source.HyperlinkKind,
            Hyperlink = source.Hyperlink,
            ImageBytes = source.ImageBytes,
            ImageContentType = source.ImageContentType,
            ImageWidth = source.ImageWidth,
            ImageHeight = source.ImageHeight,
            StyleIndex = source.StyleIndex,
            Style = source.Style,
            Region = source.Region
        };
    }

    private sealed class CollectedObject
    {
        public string Name { get; init; } = string.Empty;

        public string Kind { get; init; } = string.Empty;

        public string ParentBandName { get; init; } = string.Empty;

        public float Left { get; init; }

        public float Top { get; init; }

        public float Width { get; init; }

        public float Height { get; init; }

        public bool IsText { get; init; }

        public string Text { get; init; } = string.Empty;

        public string Bookmark { get; init; } = string.Empty;

        public HyperlinkKind? HyperlinkKind { get; init; }

        public string Hyperlink { get; init; } = string.Empty;

        public byte[]? ImageBytes { get; init; }

        public string ImageContentType { get; init; } = string.Empty;

        public float ImageWidth { get; init; }

        public float ImageHeight { get; init; }

        public int StyleIndex { get; init; }

        public DocxMatrixStyle Style { get; init; } = DocxMatrixStyle.Empty;

        public DocxPageRegion Region { get; init; }
    }
}
