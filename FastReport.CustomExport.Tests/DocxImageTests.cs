using FastReport.Export.Custom;
using System.Drawing;
using System.IO;
using System.Linq;
using Xunit;

namespace FastReport.CustomExport.Tests;

public class DocxImageTests
{
    [Fact]
    public void ExportedDocxShouldWriteBodyPicturesAsMainDocumentImages()
    {
        using Report report = CreateBodyImageReport();
        using MemoryStream stream = new MemoryStream();
        using DocxExport export = new DocxExport();

        report.Export(export, stream);

        string[] bodyTexts = DocxLayoutReader.ReadParagraphLayouts(stream)
            .Select(x => x.Text)
            .ToArray();

        Assert.Contains("Body text", bodyTexts);
        Assert.True(DocxLayoutReader.ReadBodyImageCount(stream) > 0);
    }

    [Fact]
    public void ExportedDocxShouldWriteHeaderAndFooterPicturesToTheirOwnParts()
    {
        using Report report = CreateHeaderFooterImageReport();
        using MemoryStream stream = new MemoryStream();
        using DocxExport export = new DocxExport();

        report.Export(export, stream);

        Assert.True(DocxLayoutReader.ReadHeaderImageCount(stream) > 0);
        Assert.True(DocxLayoutReader.ReadFooterImageCount(stream) > 0);
        Assert.Equal(0, DocxLayoutReader.ReadBodyImageCount(stream));
    }

    private static Report CreateBodyImageReport()
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
            Name = "BodyText",
            Text = "Body text",
            Left = 10,
            Top = 10,
            Width = 120,
            Height = 20
        });

        band.Objects.Add(new PictureObject
        {
            Name = "Picture1",
            Left = 10,
            Top = 40,
            Width = 60,
            Height = 30,
            Image = CreateBitmap(Color.OrangeRed)
        });

        page.ReportTitle = band;
        report.Pages.Add(page);
        report.Prepare();
        return report;
    }

    private static Report CreateHeaderFooterImageReport()
    {
        Report report = new Report();
        ReportPage page = new ReportPage();

        PageHeaderBand header = new PageHeaderBand
        {
            Name = "PageHeader1",
            Height = 32
        };
        header.Objects.Add(new PictureObject
        {
            Name = "HeaderPicture",
            Left = 8,
            Top = 0,
            Width = 40,
            Height = 20,
            Image = CreateBitmap(Color.ForestGreen)
        });

        ReportTitleBand body = new ReportTitleBand
        {
            Name = "ReportTitle1",
            Height = 80
        };
        body.Objects.Add(new TextObject
        {
            Name = "BodyText",
            Text = "Body text",
            Left = 10,
            Top = 10,
            Width = 100,
            Height = 20
        });

        PageFooterBand footer = new PageFooterBand
        {
            Name = "PageFooter1",
            Height = 32
        };
        footer.Objects.Add(new PictureObject
        {
            Name = "FooterPicture",
            Left = 8,
            Top = 0,
            Width = 40,
            Height = 20,
            Image = CreateBitmap(Color.SteelBlue)
        });

        page.PageHeader = header;
        page.ReportTitle = body;
        page.PageFooter = footer;
        report.Pages.Add(page);
        report.Prepare();
        return report;
    }

    private static Bitmap CreateBitmap(Color color)
    {
        Bitmap bitmap = new Bitmap(12, 8);
        using Graphics graphics = Graphics.FromImage(bitmap);
        graphics.Clear(color);
        return bitmap;
    }
}
