using DocumentFormat.OpenXml.Wordprocessing;
using FastReport.Export.Custom;
using Xunit;
using DrawingColor = System.Drawing.Color;

namespace FastReport.CustomExport.Tests;

public class DocxPageLayoutTests
{
    [Fact]
    public void ExportedDocxShouldPreservePageSizeAndOrientation()
    {
        using Report report = CreateSinglePageReport(
            width: 210,
            height: 297,
            landscape: false);
        ReportPage preparedPage = report.PreparedPages.GetPage(0);

        using MemoryStream stream = new MemoryStream();
        DocxExport export = new DocxExport();
        report.Export(export, stream);

        DocxPageLayout layout = DocxLayoutReader.ReadPageLayout(stream);

        Assert.Equal(Math.Max(MillimetersToTwips(210), PixelsToTwips(preparedPage.WidthInPixels)), layout.Width);
        Assert.Equal(Math.Max(MillimetersToTwips(297), PixelsToTwips(preparedPage.HeightInPixels)), layout.Height);
        Assert.Equal(PageOrientationValues.Portrait, layout.Orientation);
    }

    [Fact]
    public void ExportedDocxShouldPreserveLandscapeOrientation()
    {
        using Report report = CreateSinglePageReport(
            width: 210,
            height: 297,
            landscape: true);
        ReportPage preparedPage = report.PreparedPages.GetPage(0);

        using MemoryStream stream = new MemoryStream();
        DocxExport export = new DocxExport();
        report.Export(export, stream);

        DocxPageLayout layout = DocxLayoutReader.ReadPageLayout(stream);

        Assert.Equal(Math.Max(MillimetersToTwips(297), PixelsToTwips(preparedPage.WidthInPixels)), layout.Width);
        Assert.Equal(Math.Max(MillimetersToTwips(210), PixelsToTwips(preparedPage.HeightInPixels)), layout.Height);
        Assert.Equal(PageOrientationValues.Landscape, layout.Orientation);
    }

    [Fact]
    public void ExportedDocxShouldUseEffectivePreparedSizeForUnlimitedPages()
    {
        using Report report = CreateUnlimitedPageReport();
        ReportPage sourcePage = (ReportPage)report.Pages[0];
        ReportPage preparedPage = report.PreparedPages.GetPage(0);

        using MemoryStream stream = new MemoryStream();
        using DocxExport export = new DocxExport();
        report.Export(export, stream);

        DocxPageLayout layout = DocxLayoutReader.ReadPageLayout(stream);
        uint expectedWidth = Math.Max(
            sourcePage.Landscape
                ? MillimetersToTwips(Math.Max(sourcePage.PaperWidth, sourcePage.PaperHeight))
                : MillimetersToTwips(sourcePage.PaperWidth),
            PixelsToTwips(preparedPage.WidthInPixels));
        uint expectedHeight = Math.Max(
            sourcePage.Landscape
                ? MillimetersToTwips(Math.Min(sourcePage.PaperWidth, sourcePage.PaperHeight))
                : MillimetersToTwips(sourcePage.PaperHeight),
            PixelsToTwips(preparedPage.HeightInPixels));

        Assert.Equal(expectedWidth, layout.Width);
        Assert.Equal(expectedHeight, layout.Height);
        Assert.True(layout.Width > MillimetersToTwips(sourcePage.PaperWidth));
        Assert.True(layout.Height > MillimetersToTwips(sourcePage.PaperHeight));
    }

    [Fact]
    public void ExportedDocxShouldUsePreparedSizeExactlyForShrinkingUnlimitedPages()
    {
        using Report report = CreateShrinkingUnlimitedPageReport();
        ReportPage preparedPage = report.PreparedPages.GetPage(0);

        using MemoryStream stream = new MemoryStream();
        using DocxExport export = new DocxExport();
        report.Export(export, stream);

        DocxPageLayout layout = DocxLayoutReader.ReadPageLayout(stream);

        Assert.Equal(PixelsToTwips(preparedPage.WidthInPixels), layout.Width);
        Assert.Equal(PixelsToTwips(preparedPage.HeightInPixels), layout.Height);
        Assert.True(layout.Height < MillimetersToTwips(210));
    }

    [Fact]
    public void ExportedDocxShouldLeaveVerticalHeadroomWhenContentMatchesPrintableHeight()
    {
        using Report report = CreateNearPrintableHeightReport();
        ReportPage preparedPage = report.PreparedPages.GetPage(0);

        using MemoryStream stream = new MemoryStream();
        using DocxExport export = new DocxExport();
        report.Export(export, stream);

        DocxParagraphLayout[] paragraphs = DocxLayoutReader.ReadParagraphLayouts(stream).ToArray();
        int bottom = paragraphs.Max(x => x.FrameY + x.FrameHeight);
        uint printableHeight = PixelsToTwips(
            preparedPage.HeightInPixels
            - MillimetersToPixels(preparedPage.TopMargin)
            - MillimetersToPixels(preparedPage.BottomMargin));

        Assert.True(bottom < (int)printableHeight);
    }

    private static Report CreateSinglePageReport(float width, float height, bool landscape)
    {
        Report report = new Report();
        ReportPage page = new ReportPage
        {
            PaperWidth = width,
            PaperHeight = height
        };

        if (landscape)
            page.Landscape = true;

        ReportTitleBand band = new ReportTitleBand
        {
            Name = "ReportTitle1",
            Height = 40
        };

        TextObject text = new TextObject
        {
            Name = "Text1",
            Text = "Layout probe",
            Width = 300,
            Height = 20
        };

        band.Objects.Add(text);
        page.ReportTitle = band;
        report.Pages.Add(page);
        report.Prepare();
        return report;
    }

    private static Report CreateUnlimitedPageReport()
    {
        Report report = new Report();
        ReportPage page = new ReportPage
        {
            PaperWidth = 210,
            PaperHeight = 297,
            UnlimitedWidth = true,
            UnlimitedHeight = true
        };

        ReportTitleBand band = new ReportTitleBand
        {
            Name = "ReportTitle1",
            Height = 1100
        };

        TextObject anchor = new TextObject
        {
            Name = "Text1",
            Text = "Anchor",
            Left = 0,
            Top = 0,
            Width = 120,
            Height = 20
        };

        TextObject farAway = new TextObject
        {
            Name = "Text2",
            Text = "Far away",
            Left = 1500,
            Top = 950,
            Width = 240,
            Height = 40
        };

        band.Objects.Add(anchor);
        band.Objects.Add(farAway);
        page.ReportTitle = band;
        report.Pages.Add(page);
        report.Prepare();
        return report;
    }

    private static Report CreateShrinkingUnlimitedPageReport()
    {
        Report report = new Report();
        ReportPage page = new ReportPage
        {
            PaperWidth = 297,
            PaperHeight = 210,
            Landscape = true,
            UnlimitedWidth = true,
            UnlimitedHeight = true
        };

        ReportTitleBand band = new ReportTitleBand
        {
            Name = "ReportTitle1",
            Height = 140
        };

        band.Objects.Add(new TextObject
        {
            Name = "Text1",
            Text = "Short unlimited page",
            Left = 0,
            Top = 0,
            Width = 800,
            Height = 20
        });

        page.ReportTitle = band;
        report.Pages.Add(page);
        report.Prepare();
        return report;
    }

    private static Report CreateNearPrintableHeightReport()
    {
        Report report = new Report();
        ReportPage page = new ReportPage
        {
            PaperWidth = 210,
            PaperHeight = 297,
            LeftMargin = 10,
            TopMargin = 10,
            RightMargin = 10,
            BottomMargin = 10
        };

        float printableHeight = TwipsToPixels(
            MillimetersToTwips(page.PaperHeight)
            - MillimetersToTwips(page.TopMargin)
            - MillimetersToTwips(page.BottomMargin));

        ReportTitleBand band = new ReportTitleBand
        {
            Name = "ReportTitle1",
            Height = printableHeight
        };

        band.Objects.Add(new TextObject
        {
            Name = "Text1",
            Text = "Boundary probe",
            Left = 0,
            Top = 0,
            Width = 300,
            Height = 20
        });

        band.Fill = new SolidFill(DrawingColor.FromArgb(220, 220, 220));

        page.ReportTitle = band;
        report.Pages.Add(page);
        report.Prepare();
        return report;
    }

    private static uint MillimetersToTwips(float millimeters)
    {
        return (uint)Math.Round(millimeters * 1440d / 25.4d, MidpointRounding.AwayFromZero);
    }

    private static uint PixelsToTwips(float pixels)
    {
        return (uint)Math.Round(pixels * 1440d / 96d, MidpointRounding.AwayFromZero);
    }

    private static int MillimetersToPixels(float millimeters)
    {
        return (int)Math.Round(millimeters * 96d / 25.4d, MidpointRounding.AwayFromZero);
    }

    private static float TwipsToPixels(uint twips)
    {
        return (float)(twips * 96d / 1440d);
    }
}
