using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Text.RegularExpressions;
using WordTable = DocumentFormat.OpenXml.Wordprocessing.Table;

namespace FastReport.CustomExport.Tests;

internal static class DocxLayoutReader
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace V = "urn:schemas-microsoft-com:vml";
    private static readonly Regex StyleValueRegex = new Regex(@"(?<name>margin-left|margin-top|width|height):(?<value>[-0-9.]+)pt", RegexOptions.Compiled);

    public static DocxMargins ReadMargins(Stream docxStream)
    {
        SectionProperties sectionProperties = ReadSectionProperties(docxStream);
        PageMargin? pageMargin = sectionProperties.GetFirstChild<PageMargin>();
        if (pageMargin == null)
            throw new InvalidOperationException("DOCX does not contain page margin information.");

        return new DocxMargins
        {
            Left = pageMargin.Left?.Value ?? 0,
            Top = (UInt32)(pageMargin.Top?.Value ?? 0),
            Right = pageMargin.Right?.Value ?? 0,
            Bottom = (UInt32)(pageMargin.Bottom?.Value ?? 0)
        };
    }

    public static DocxPageLayout ReadPageLayout(Stream docxStream)
    {
        SectionProperties sectionProperties = ReadSectionProperties(docxStream);
        PageSize? pageSize = sectionProperties.GetFirstChild<PageSize>();
        if (pageSize == null)
            throw new InvalidOperationException("DOCX does not contain page size information.");

        return new DocxPageLayout
        {
            Width = pageSize.Width?.Value ?? 0,
            Height = pageSize.Height?.Value ?? 0,
            Orientation = pageSize.Orient?.Value ?? PageOrientationValues.Portrait
        };
    }

    public static IReadOnlyList<DocxParagraphLayout> ReadParagraphLayouts(Stream docxStream)
    {
        if (docxStream.CanSeek)
            docxStream.Position = 0;

        using WordprocessingDocument document = WordprocessingDocument.Open(docxStream, false);
        Body? body = document.MainDocumentPart?.Document?.Body;
        if (body == null)
            return Array.Empty<DocxParagraphLayout>();

        IReadOnlyList<DocxParagraphLayout> tableLayouts = body.Elements<WordTable>()
            .SelectMany(ReadTableLayouts)
            .ToList();

        if (tableLayouts.Count > 0)
            return tableLayouts;

        using Stream xmlStream = document.MainDocumentPart?.GetStream()
            ?? throw new InvalidOperationException("DOCX does not contain a main document part.");
        XDocument xml = XDocument.Load(xmlStream);
        if (xml.Root == null)
            return Array.Empty<DocxParagraphLayout>();

        return xml.Descendants(V + "shape")
            .Select(ReadShapeLayout)
            .ToList();
    }

    public static IReadOnlyList<DocxRunLayout> ReadRunLayouts(Stream docxStream)
    {
        if (docxStream.CanSeek)
            docxStream.Position = 0;

        using WordprocessingDocument document = WordprocessingDocument.Open(docxStream, false);
        Body? body = document.MainDocumentPart?.Document?.Body;
        if (body == null)
            return Array.Empty<DocxRunLayout>();

        return body.Descendants<Run>()
            .Select(ReadRunLayout)
            .ToList();
    }

    public static IReadOnlyList<string> ReadHeaderTexts(Stream docxStream)
    {
        return ReadPartTexts(
            docxStream,
            document => document.MainDocumentPart?.HeaderParts.Select(part => part.Header));
    }

    public static IReadOnlyList<string> ReadFooterTexts(Stream docxStream)
    {
        return ReadPartTexts(
            docxStream,
            document => document.MainDocumentPart?.FooterParts.Select(part => part.Footer));
    }

    public static int ReadBodyImageCount(Stream docxStream)
    {
        if (docxStream.CanSeek)
            docxStream.Position = 0;

        using WordprocessingDocument document = WordprocessingDocument.Open(docxStream, false);
        return document.MainDocumentPart?.ImageParts.Count() ?? 0;
    }

    public static int ReadHeaderImageCount(Stream docxStream)
    {
        if (docxStream.CanSeek)
            docxStream.Position = 0;

        using WordprocessingDocument document = WordprocessingDocument.Open(docxStream, false);
        return document.MainDocumentPart?.HeaderParts.Sum(part => part.ImageParts.Count()) ?? 0;
    }

    public static int ReadFooterImageCount(Stream docxStream)
    {
        if (docxStream.CanSeek)
            docxStream.Position = 0;

        using WordprocessingDocument document = WordprocessingDocument.Open(docxStream, false);
        return document.MainDocumentPart?.FooterParts.Sum(part => part.ImageParts.Count()) ?? 0;
    }

    private static SectionProperties ReadSectionProperties(Stream docxStream)
    {
        if (docxStream.CanSeek)
            docxStream.Position = 0;

        using WordprocessingDocument document = WordprocessingDocument.Open(docxStream, false);
        SectionProperties? sectionProperties = document.MainDocumentPart?
            .Document?
            .Body?
            .GetFirstChild<SectionProperties>();
        if (sectionProperties == null)
            throw new InvalidOperationException("DOCX does not contain section properties.");

        return sectionProperties.CloneNode(true) as SectionProperties
            ?? throw new InvalidOperationException("DOCX section properties could not be cloned.");
    }

    private static IReadOnlyList<string> ReadPartTexts(
        Stream docxStream,
        Func<WordprocessingDocument, IEnumerable<OpenXmlPartRootElement?>?> selector)
    {
        if (docxStream.CanSeek)
            docxStream.Position = 0;

        using WordprocessingDocument document = WordprocessingDocument.Open(docxStream, false);
        IEnumerable<string> texts = selector(document)?
            .Where(part => part != null)
            .SelectMany(part => part!.Descendants<Text>())
            .Select(text => text.Text)
            .Where(text => !string.IsNullOrWhiteSpace(text))
            ?? Enumerable.Empty<string>();

        return texts.ToList();
    }

    private static DocxParagraphLayout ReadShapeLayout(XElement shape)
    {
        XElement? paragraph = shape
            .Descendants(W + "txbxContent")
            .Descendants(W + "p")
            .FirstOrDefault();
        string style = shape.Attribute("style")?.Value ?? string.Empty;
        Dictionary<string, double> styleValues = ParseStyleValues(style);
        string fillColor = (shape.Attribute("fillcolor")?.Value ?? string.Empty).TrimStart('#');
        string strokeColor = (shape.Attribute("strokecolor")?.Value ?? string.Empty).TrimStart('#');
        XElement? paragraphProperties = paragraph?.Element(W + "pPr");

        return new DocxParagraphLayout
        {
            Text = string.Concat(paragraph?.Descendants(W + "t").Select(x => x.Value) ?? Enumerable.Empty<string>()),
            LeftIndent = 0,
            RightIndent = 0,
            BeforeSpacing = 0,
            LineSpacing = ParseInt(paragraphProperties?.Element(W + "spacing")?.Attribute(W + "line")?.Value),
            LineSpacingRule = ParseLineRule(paragraphProperties?.Element(W + "spacing")?.Attribute(W + "lineRule")?.Value),
            Justification = ParseJustification(paragraphProperties),
            TabCount = paragraph?.Descendants(W + "tab").Count() ?? 0,
            LineBreakCount = paragraph?.Descendants(W + "br").Count() ?? 0,
            FrameX = PointsToTwips(styleValues.TryGetValue("margin-left", out double left) ? left : 0d),
            FrameY = PointsToTwips(styleValues.TryGetValue("margin-top", out double top) ? top : 0d),
            FrameWidth = PointsToTwips(styleValues.TryGetValue("width", out double width) ? width : 0d),
            FrameHeight = PointsToTwips(styleValues.TryGetValue("height", out double height) ? height : 0d),
            FillColor = fillColor,
            BorderTopColor = strokeColor,
            BorderLeftColor = strokeColor,
            BorderBottomColor = strokeColor,
            BorderRightColor = strokeColor
        };
    }

    private static IReadOnlyList<DocxParagraphLayout> ReadTableLayouts(WordTable table)
    {
        int[] columnWidths = table.GetFirstChild<TableGrid>()?
            .Elements<GridColumn>()
            .Select(x => Int32.TryParse(x.Width?.Value, out int width) ? width : 0)
            .ToArray() ?? Array.Empty<int>();

        List<DocxParagraphLayout> layouts = new();
        Dictionary<int, ActiveVerticalMerge> activeMerges = new();
        int currentY = 0;

        foreach (TableRow row in table.Elements<TableRow>())
        {
            int rowHeight = (int)(row.GetFirstChild<TableRowProperties>()?.GetFirstChild<TableRowHeight>()?.Val?.Value ?? 0U);
            int currentX = 0;
            int columnIndex = 0;

            foreach (TableCell cell in row.Elements<TableCell>())
            {
                while (activeMerges.TryGetValue(columnIndex, out ActiveVerticalMerge? active) && active.RemainingRows == 0)
                    activeMerges.Remove(columnIndex);

                TableCellProperties? properties = cell.GetFirstChild<TableCellProperties>();
                int gridSpan = properties?.GetFirstChild<GridSpan>()?.Val?.Value ?? 1;
                VerticalMerge? verticalMerge = properties?.GetFirstChild<VerticalMerge>();

                if (verticalMerge != null && verticalMerge.Val == null)
                {
                    if (activeMerges.TryGetValue(columnIndex, out ActiveVerticalMerge? continued))
                    {
                        continued.RemainingRows--;
                        activeMerges[columnIndex] = continued;
                    }

                    currentX += Sum(columnWidths, columnIndex, gridSpan);
                    columnIndex += gridSpan;
                    continue;
                }

                int width = Sum(columnWidths, columnIndex, gridSpan);
                int height = rowHeight;
                if (verticalMerge?.Val?.Value == MergedCellValues.Restart)
                {
                    int extraRows = CountContinuationRows(row, columnIndex, gridSpan);
                    if (extraRows > 0)
                    {
                        activeMerges[columnIndex] = new ActiveVerticalMerge(extraRows);
                        height = rowHeight + SumContinuationHeights(row, extraRows);
                    }
                }

                string text = string.Concat(cell.Descendants<Text>().Select(x => x.Text));
                Paragraph? paragraph = cell.Elements<Paragraph>().FirstOrDefault();
                ParagraphProperties? paragraphProperties = paragraph?.ParagraphProperties;
                TableCellBorders? borders = properties?.GetFirstChild<TableCellBorders>();
                Shading? shading = properties?.GetFirstChild<Shading>();

                layouts.Add(new DocxParagraphLayout
                {
                    Text = text,
                    LeftIndent = 0,
                    RightIndent = 0,
                    BeforeSpacing = 0,
                    LineSpacing = Int32.TryParse(paragraphProperties?.GetFirstChild<SpacingBetweenLines>()?.Line?.Value, out int lineSpacing) ? lineSpacing : 0,
                    LineSpacingRule = paragraphProperties?.GetFirstChild<SpacingBetweenLines>()?.LineRule?.Value,
                    Justification = paragraphProperties?.GetFirstChild<Justification>()?.Val?.Value,
                    TabCount = paragraph?.Descendants<TabChar>().Count() ?? 0,
                    LineBreakCount = paragraph?.Descendants<Break>().Count() ?? 0,
                    FrameX = currentX,
                    FrameY = currentY,
                    FrameWidth = width,
                    FrameHeight = height,
                    FillColor = shading?.Fill?.Value ?? string.Empty,
                    BorderTopColor = borders?.TopBorder?.Color?.Value ?? string.Empty,
                    BorderLeftColor = borders?.LeftBorder?.Color?.Value ?? string.Empty,
                    BorderBottomColor = borders?.BottomBorder?.Color?.Value ?? string.Empty,
                    BorderRightColor = borders?.RightBorder?.Color?.Value ?? string.Empty
                });

                currentX += width;
                columnIndex += gridSpan;
            }

            currentY += rowHeight;
            foreach (int key in activeMerges.Keys.ToList())
            {
                ActiveVerticalMerge merge = activeMerges[key];
                if (merge.RemainingRows > 0)
                {
                    merge.RemainingRows--;
                    activeMerges[key] = merge;
                }
            }
        }

        return layouts.Where(x => !string.IsNullOrWhiteSpace(x.Text)).ToList();
    }

    private static int CountContinuationRows(TableRow startRow, int columnIndex, int gridSpan)
    {
        int count = 0;
        TableRow? row = startRow.NextSibling<TableRow>();
        while (row != null)
        {
            TableCell? cell = GetCellAtColumn(row, columnIndex);
            if (cell?.GetFirstChild<TableCellProperties>()?.GetFirstChild<VerticalMerge>()?.Val != null)
                break;

            if (cell?.GetFirstChild<TableCellProperties>()?.GetFirstChild<VerticalMerge>() == null)
                break;

            count++;
            row = row.NextSibling<TableRow>();
        }

        return count;
    }

    private static int SumContinuationHeights(TableRow startRow, int count)
    {
        int sum = 0;
        TableRow? row = startRow.NextSibling<TableRow>();
        while (count > 0 && row != null)
        {
            sum += (int)(row.GetFirstChild<TableRowProperties>()?.GetFirstChild<TableRowHeight>()?.Val?.Value ?? 0U);
            count--;
            row = row.NextSibling<TableRow>();
        }

        return sum;
    }

    private static TableCell? GetCellAtColumn(TableRow row, int targetColumn)
    {
        int currentColumn = 0;
        foreach (TableCell cell in row.Elements<TableCell>())
        {
            int span = cell.GetFirstChild<TableCellProperties>()?.GetFirstChild<GridSpan>()?.Val?.Value ?? 1;
            if (currentColumn == targetColumn)
                return cell;

            currentColumn += span;
        }

        return null;
    }

    private static int Sum(int[] widths, int start, int count)
    {
        int total = 0;
        for (int i = start; i < Math.Min(widths.Length, start + count); i++)
            total += widths[i];

        return total;
    }

    private static Dictionary<string, double> ParseStyleValues(string style)
    {
        Dictionary<string, double> values = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        foreach (Match match in StyleValueRegex.Matches(style))
        {
            values[match.Groups["name"].Value] = double.Parse(match.Groups["value"].Value, System.Globalization.CultureInfo.InvariantCulture);
        }

        return values;
    }

    private static JustificationValues? ParseJustification(XElement? paragraphProperties)
    {
        string? value = paragraphProperties?
            .Element(W + "jc")?
            .Attribute(W + "val")?
            .Value;

        return value switch
        {
            "left" => JustificationValues.Left,
            "center" => JustificationValues.Center,
            "right" => JustificationValues.Right,
            "both" => JustificationValues.Both,
            _ => null
        };
    }

    private static int PointsToTwips(double points)
    {
        return (int)Math.Round(points * 20d, MidpointRounding.AwayFromZero);
    }

    private static int ParseInt(string? value)
    {
        return Int32.TryParse(value, out int result) ? result : 0;
    }

    private static LineSpacingRuleValues? ParseLineRule(string? value)
    {
        return value switch
        {
            "auto" => LineSpacingRuleValues.Auto,
            "atLeast" => LineSpacingRuleValues.AtLeast,
            "exact" => LineSpacingRuleValues.Exact,
            _ => null
        };
    }

    private static DocxRunLayout ReadRunLayout(Run run)
    {
        RunProperties? properties = run.RunProperties;
        return new DocxRunLayout
        {
            Text = string.Concat(run.Descendants<Text>().Select(x => x.Text)),
            FontName = properties?.RunFonts?.Ascii?.Value ?? string.Empty,
            FontSize = Int32.TryParse(properties?.FontSize?.Val?.Value, out int fontSize) ? fontSize : 0,
            Bold = properties?.Bold != null,
            Italic = properties?.Italic != null,
            Underline = properties?.Underline?.Val?.Value,
            Color = properties?.Color?.Val?.Value ?? string.Empty
        };
    }
}

internal sealed class DocxMargins
{
    public UInt32 Left { get; init; }

    public UInt32 Top { get; init; }

    public UInt32 Right { get; init; }

    public UInt32 Bottom { get; init; }
}

internal sealed class DocxPageLayout
{
    public UInt32 Width { get; init; }

    public UInt32 Height { get; init; }

    public PageOrientationValues Orientation { get; init; }
}

internal sealed class DocxParagraphLayout
{
    public string Text { get; init; } = string.Empty;

    public Int32 LeftIndent { get; init; }

    public Int32 RightIndent { get; init; }

    public Int32 BeforeSpacing { get; init; }

    public Int32 LineSpacing { get; init; }

    public LineSpacingRuleValues? LineSpacingRule { get; init; }

    public JustificationValues? Justification { get; init; }

    public Int32 TabCount { get; init; }

    public Int32 LineBreakCount { get; init; }

    public Int32 FrameX { get; init; }

    public Int32 FrameY { get; init; }

    public Int32 FrameWidth { get; init; }

    public Int32 FrameHeight { get; init; }

    public string FillColor { get; init; } = string.Empty;

    public string BorderTopColor { get; init; } = string.Empty;

    public string BorderLeftColor { get; init; } = string.Empty;

    public string BorderBottomColor { get; init; } = string.Empty;

    public string BorderRightColor { get; init; } = string.Empty;
}

internal sealed class DocxRunLayout
{
    public string Text { get; init; } = string.Empty;

    public string FontName { get; init; } = string.Empty;

    public Int32 FontSize { get; init; }

    public bool Bold { get; init; }

    public bool Italic { get; init; }

    public UnderlineValues? Underline { get; init; }

    public string Color { get; init; } = string.Empty;
}

internal sealed class ActiveVerticalMerge
{
    public ActiveVerticalMerge(int remainingRows)
    {
        RemainingRows = remainingRows;
    }

    public int RemainingRows { get; set; }
}
