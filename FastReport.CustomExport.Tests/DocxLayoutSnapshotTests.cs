using FastReport.Export.Custom;
using System.Linq;
using Xunit;

namespace FastReport.CustomExport.Tests;

public class DocxLayoutSnapshotTests
{
    [Fact]
    public void SnapshotBuilderShouldCaptureBandsAndTextObjectsFromPreparedPage()
    {
        using Report report = CreateStructuredReport();
        using ReportPage page = report.PreparedPages.GetPage(0);

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshotBuilder.Build(page);

        Assert.Equal(page.WidthInPixels, snapshot.PageWidth);
        Assert.Equal(page.HeightInPixels, snapshot.PageHeight);
        Assert.Contains(snapshot.Bands, x => x.Name == "ReportTitle1");
        Assert.Contains(snapshot.Bands, x => x.Name == "PageFooter1");
        Assert.Contains(snapshot.Texts, x => x.Name == "TitleText");
        Assert.Contains(snapshot.Texts, x => x.Name == "FooterText");
    }

    [Fact]
    public void SnapshotBuilderShouldKeepAbsoluteGeometryAndBandOwnership()
    {
        using Report report = CreateStructuredReport();
        using ReportPage page = report.PreparedPages.GetPage(0);

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshotBuilder.Build(page);
        DocxTextSnapshot title = snapshot.Texts.Single(x => x.Name == "TitleText");
        DocxTextSnapshot footer = snapshot.Texts.Single(x => x.Name == "FooterText");

        Assert.Equal("ReportTitle1", title.ParentBandName);
        Assert.Equal("PageFooter1", footer.ParentBandName);
        Assert.True(footer.Top > title.Top);
        Assert.True(title.Width > 0);
        Assert.True(footer.Height > 0);
    }

    private static Report CreateStructuredReport()
    {
        Report report = new Report();
        ReportPage page = new ReportPage
        {
            PaperWidth = 210,
            PaperHeight = 297
        };

        ReportTitleBand titleBand = new ReportTitleBand
        {
            Name = "ReportTitle1",
            Height = 70
        };
        titleBand.Objects.Add(new TextObject
        {
            Name = "TitleText",
            Text = "Title",
            Left = 20,
            Top = 12,
            Width = 180,
            Height = 20
        });

        DataBand dataBand = new DataBand
        {
            Name = "Data1",
            Height = 120
        };
        dataBand.Objects.Add(new TextObject
        {
            Name = "BodyText",
            Text = "Body",
            Left = 30,
            Top = 10,
            Width = 160,
            Height = 20
        });

        PageFooterBand footerBand = new PageFooterBand
        {
            Name = "PageFooter1",
            Height = 28
        };
        footerBand.Objects.Add(new TextObject
        {
            Name = "FooterText",
            Text = "Footer",
            Left = 10,
            Top = 0,
            Width = 100,
            Height = 20
        });

        page.ReportTitle = titleBand;
        page.Bands.Add(dataBand);
        page.PageFooter = footerBand;
        report.Pages.Add(page);
        report.Prepare();
        return report;
    }
}
