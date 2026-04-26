using DocumentFormat.OpenXml.Wordprocessing;
using FastReport.Export.Custom;
using System.Drawing;
using System.Linq;
using Xunit;
using DrawingColor = System.Drawing.Color;
using DrawingFont = System.Drawing.Font;

namespace FastReport.CustomExport.Tests;

public class DocxFontStyleTests
{
    [Fact]
    public void ExportedDocxShouldPreserveBasicFontFormatting()
    {
        using Report report = CreateStyledReport();
        using MemoryStream stream = new MemoryStream();
        DocxExport export = new DocxExport();

        report.Export(export, stream);

        DocxRunLayout run = DocxLayoutReader.ReadRunLayouts(stream)
            .First(x => x.Text == "Styled text" && !string.IsNullOrEmpty(x.FontName));

        Assert.Equal("Verdana", run.FontName);
        Assert.Equal(28, run.FontSize);
        Assert.True(run.Bold);
        Assert.True(run.Italic);
        Assert.Equal(UnderlineValues.Single, run.Underline);
        Assert.Equal("0000FF", run.Color);
    }

    private static Report CreateStyledReport()
    {
        Report report = new Report();
        ReportPage page = new ReportPage();
        ReportTitleBand band = new ReportTitleBand
        {
            Name = "ReportTitle1",
            Height = 100
        };

        TextObject text = new TextObject
        {
            Name = "Text1",
            Text = "Styled text",
            Width = 200,
            Height = 30,
            Font = new DrawingFont("Verdana", 14, FontStyle.Bold | FontStyle.Italic),
            Underlines = true,
            TextColor = DrawingColor.Blue
        };

        band.Objects.Add(text);
        page.ReportTitle = band;
        report.Pages.Add(page);
        report.Prepare();
        return report;
    }
}
