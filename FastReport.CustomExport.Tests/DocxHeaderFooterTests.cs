using FastReport.Export.Custom;
using System.IO;
using System.Linq;
using Xunit;

namespace FastReport.CustomExport.Tests;

public class DocxHeaderFooterTests
{
    [Fact]
    public void ExportedDocxShouldMovePageBandsIntoHeaderAndFooterParts()
    {
        using Report report = CreateHeaderFooterReport();
        using MemoryStream stream = new MemoryStream();
        using DocxExport export = new DocxExport();

        report.Export(export, stream);

        string[] bodyTexts = DocxLayoutReader.ReadParagraphLayouts(stream)
            .Select(x => x.Text)
            .ToArray();
        string[] headerTexts = DocxLayoutReader.ReadHeaderTexts(stream).ToArray();
        string[] footerTexts = DocxLayoutReader.ReadFooterTexts(stream).ToArray();

        Assert.Contains("Header text", headerTexts);
        Assert.Contains("Footer text", footerTexts);
        Assert.Contains("Body text", bodyTexts);
        Assert.DoesNotContain("Header text", bodyTexts);
        Assert.DoesNotContain("Footer text", bodyTexts);
    }

    private static Report CreateHeaderFooterReport()
    {
        Report report = new Report();
        ReportPage page = new ReportPage();

        PageHeaderBand headerBand = new PageHeaderBand
        {
            Name = "PageHeader1",
            Height = 28
        };
        headerBand.Objects.Add(new TextObject
        {
            Name = "HeaderText",
            Text = "Header text",
            Left = 10,
            Top = 0,
            Width = 120,
            Height = 20
        });

        ReportTitleBand titleBand = new ReportTitleBand
        {
            Name = "ReportTitle1",
            Height = 120
        };
        titleBand.Objects.Add(new TextObject
        {
            Name = "BodyText",
            Text = "Body text",
            Left = 20,
            Top = 20,
            Width = 140,
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
            Text = "Footer text",
            Left = 10,
            Top = 0,
            Width = 120,
            Height = 20
        });

        page.PageHeader = headerBand;
        page.ReportTitle = titleBand;
        page.PageFooter = footerBand;
        report.Pages.Add(page);
        report.Prepare();
        return report;
    }
}
