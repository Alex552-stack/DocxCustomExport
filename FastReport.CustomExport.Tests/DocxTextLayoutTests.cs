using DocumentFormat.OpenXml.Wordprocessing;
using FastReport.Export.Custom;
using System.Drawing;
using System.Linq;
using Xunit;
using DrawingFont = System.Drawing.Font;

namespace FastReport.CustomExport.Tests;

public class DocxTextLayoutTests
{
    [Fact]
    public void ExportedDocxShouldMapTextObjectPositioningToParagraphLayout()
    {
        using Report report = CreateStructuredReport();
        using MemoryStream stream = new MemoryStream();
        DocxExport export = new DocxExport();

        report.Export(export, stream);

        DocxParagraphLayout[] paragraphs = DocxLayoutReader.ReadParagraphLayouts(stream)
            .Where(x => x.Text is "First row" or "Second row")
            .ToArray();

        Assert.Equal(2, paragraphs.Length);

        Assert.Equal("First row", paragraphs[0].Text);
        Assert.Equal(300, paragraphs[0].FrameX);
        Assert.Equal(150, paragraphs[0].FrameY);
        Assert.Equal(2250, paragraphs[0].FrameWidth);
        Assert.Equal(300, paragraphs[0].FrameHeight);
        Assert.Equal(JustificationValues.Left, paragraphs[0].Justification);

        Assert.Equal("Second row", paragraphs[1].Text);
        Assert.Equal(600, paragraphs[1].FrameX);
        Assert.Equal(900, paragraphs[1].FrameY);
        Assert.Equal(2250, paragraphs[1].FrameWidth);
        Assert.Equal(300, paragraphs[1].FrameHeight);
        Assert.Equal(JustificationValues.Right, paragraphs[1].Justification);
    }

    [Fact]
    public void ExportedDocxShouldKeepSameRowObjectsAsIndependentFrames()
    {
        using Report report = CreateSameRowReport();
        using MemoryStream stream = new MemoryStream();
        DocxExport export = new DocxExport();

        report.Export(export, stream);

        DocxParagraphLayout[] paragraphs = DocxLayoutReader.ReadParagraphLayouts(stream)
            .Where(x => x.Text is "Left" or "Right")
            .OrderBy(x => x.FrameX)
            .ToArray();

        Assert.Equal(2, paragraphs.Length);
        Assert.Equal("Left", paragraphs[0].Text);
        Assert.Equal(150, paragraphs[0].FrameX);
        Assert.Equal("Right", paragraphs[1].Text);
        Assert.Equal(1800, paragraphs[1].FrameX);
    }

    [Fact]
    public void ExportedDocxShouldPreserveLineBreaksInsideTextObjects()
    {
        using Report report = CreateMultilineReport();
        using MemoryStream stream = new MemoryStream();
        DocxExport export = new DocxExport();

        report.Export(export, stream);

        DocxParagraphLayout paragraph = DocxLayoutReader.ReadParagraphLayouts(stream)
            .Single(x => x.Text == "Top lineBottom lineThird line");

        Assert.Equal(2, paragraph.LineBreakCount);
        Assert.True(paragraph.LineSpacing > 0);
        Assert.Equal(LineSpacingRuleValues.Exact, paragraph.LineSpacingRule);
    }

    [Fact]
    public void ExportedDocxShouldUseRenderedFontLineHeightWhenLineHeightIsImplicit()
    {
        using Report report = CreateMultilineReport();
        TextObject source = (TextObject)report.FindObject("MultiText")!;
        using MemoryStream stream = new MemoryStream();
        using DocxExport export = new DocxExport();

        report.Export(export, stream);

        DocxParagraphLayout paragraph = DocxLayoutReader.ReadParagraphLayouts(stream)
            .Single(x => x.Text == "Top lineBottom lineThird line");

        using DrawingFont font = new DrawingFont(source.Font.Name, source.Font.SizeInPoints, source.Font.Style, GraphicsUnit.Point);
        int expectedTwips = PixelsToTwips(font.GetHeight());

        Assert.Equal(expectedTwips, paragraph.LineSpacing);
        Assert.Equal(LineSpacingRuleValues.Exact, paragraph.LineSpacingRule);
    }

    [Fact]
    public void ExportedDocxShouldUseExplicitLineHeightWhenProvided()
    {
        using Report report = CreateExplicitLineHeightReport();
        using MemoryStream stream = new MemoryStream();
        using DocxExport export = new DocxExport();

        report.Export(export, stream);

        DocxParagraphLayout paragraph = DocxLayoutReader.ReadParagraphLayouts(stream)
            .Single(x => x.Text == "AlphaBetaGamma");

        Assert.Equal(PixelsToTwips(24f), paragraph.LineSpacing);
        Assert.Equal(LineSpacingRuleValues.Exact, paragraph.LineSpacingRule);
    }

    [Fact]
    public void ExportedDocxShouldCenterWithinObjectWidthNotWholePage()
    {
        using Report report = CreateCenteredObjectReport();
        using MemoryStream stream = new MemoryStream();
        DocxExport export = new DocxExport();

        report.Export(export, stream);

        DocxParagraphLayout paragraph = DocxLayoutReader.ReadParagraphLayouts(stream)
            .Single(x => x.Text == "Centered title");

        Assert.Equal(750, paragraph.FrameX);
        Assert.Equal(1950, paragraph.FrameWidth);
        Assert.Equal(JustificationValues.Center, paragraph.Justification);
    }

    private static Report CreateStructuredReport()
    {
        Report report = new Report();
        ReportPage page = new ReportPage();
        ReportTitleBand band = new ReportTitleBand
        {
            Name = "ReportTitle1",
            Height = 200
        };

        TextObject first = new TextObject
        {
            Name = "Text1",
            Text = "First row",
            Left = 20,
            Top = 10,
            Width = 150,
            Height = 20,
            HorzAlign = HorzAlign.Left
        };

        TextObject second = new TextObject
        {
            Name = "Text2",
            Text = "Second row",
            Left = 40,
            Top = 60,
            Width = 150,
            Height = 20,
            HorzAlign = HorzAlign.Right
        };

        band.Objects.Add(first);
        band.Objects.Add(second);
        page.ReportTitle = band;
        report.Pages.Add(page);
        report.Prepare();
        return report;
    }

    private static Report CreateSameRowReport()
    {
        Report report = new Report();
        ReportPage page = new ReportPage();
        ReportTitleBand band = new ReportTitleBand
        {
            Name = "ReportTitle1",
            Height = 120
        };

        band.Objects.Add(new TextObject
        {
            Name = "LeftText",
            Text = "Left",
            Left = 10,
            Top = 10,
            Width = 80,
            Height = 20
        });

        band.Objects.Add(new TextObject
        {
            Name = "RightText",
            Text = "Right",
            Left = 120,
            Top = 11,
            Width = 80,
            Height = 20
        });

        page.ReportTitle = band;
        report.Pages.Add(page);
        report.Prepare();
        return report;
    }

    private static Report CreateMultilineReport()
    {
        Report report = new Report();
        ReportPage page = new ReportPage();
        ReportTitleBand band = new ReportTitleBand
        {
            Name = "ReportTitle1",
            Height = 120
        };

        band.Objects.Add(new TextObject
        {
            Name = "MultiText",
            Text = "Top line\r\nBottom line\r\nThird line",
            Left = 10,
            Top = 10,
            Width = 120,
            Height = 60
        });

        page.ReportTitle = band;
        report.Pages.Add(page);
        report.Prepare();
        return report;
    }

    private static Report CreateExplicitLineHeightReport()
    {
        Report report = new Report();
        ReportPage page = new ReportPage();
        ReportTitleBand band = new ReportTitleBand
        {
            Name = "ReportTitle1",
            Height = 120
        };

        band.Objects.Add(new TextObject
        {
            Name = "ExplicitLineHeightText",
            Text = "Alpha\r\nBeta\r\nGamma",
            Left = 10,
            Top = 10,
            Width = 120,
            Height = 80,
            LineHeight = 24
        });

        page.ReportTitle = band;
        report.Pages.Add(page);
        report.Prepare();
        return report;
    }

    private static Report CreateCenteredObjectReport()
    {
        Report report = new Report();
        ReportPage page = new ReportPage
        {
            PaperWidth = 210,
            PaperHeight = 297,
            LeftMargin = 10,
            RightMargin = 10
        };

        ReportTitleBand band = new ReportTitleBand
        {
            Name = "ReportTitle1",
            Height = 120
        };

        band.Objects.Add(new TextObject
        {
            Name = "CenteredText",
            Text = "Centered title",
            Left = 50,
            Top = 10,
            Width = 130,
            Height = 20,
            HorzAlign = HorzAlign.Center
        });

        page.ReportTitle = band;
        report.Pages.Add(page);
        report.Prepare();
        return report;
    }

    private static int PixelsToTwips(float pixels)
    {
        return (int)System.Math.Round(pixels * 1440d / 96d, System.MidpointRounding.AwayFromZero);
    }
}
