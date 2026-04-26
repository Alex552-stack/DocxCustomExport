using FastReport.Export.Custom;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Xunit;
using DrawingColor = System.Drawing.Color;

namespace FastReport.CustomExport.Tests;

public class DocxObjectStyleTests
{
    [Fact]
    public void ExportedDocxShouldPreserveSolidFillAndBorders()
    {
        using Report report = CreateFilledBorderedReport();
        using MemoryStream stream = new MemoryStream();
        using DocxExport export = new DocxExport();

        report.Export(export, stream);

        DocxParagraphLayout paragraph = DocxLayoutReader.ReadParagraphLayouts(stream)
            .Single(x => x.Text == "Button");

        Assert.Equal("9CC2FF", paragraph.FillColor);
        Assert.Equal("336699", paragraph.BorderTopColor);
        Assert.Equal("336699", paragraph.BorderLeftColor);
        Assert.Equal("336699", paragraph.BorderBottomColor);
        Assert.Equal("336699", paragraph.BorderRightColor);
    }

    private static Report CreateFilledBorderedReport()
    {
        Report report = new Report();
        ReportPage page = new ReportPage();
        ReportTitleBand band = new ReportTitleBand
        {
            Name = "ReportTitle1",
            Height = 120
        };

        TextObject text = new TextObject
        {
            Name = "ButtonText",
            Text = "Button",
            Left = 10,
            Top = 10,
            Width = 120,
            Height = 24,
            Fill = new SolidFill(DrawingColor.FromArgb(0x9C, 0xC2, 0xFF)),
            TextColor = DrawingColor.FromArgb(0x11, 0x22, 0x33)
        };
        text.Border.Lines = BorderLines.All;
        text.Border.Color = DrawingColor.FromArgb(0x33, 0x66, 0x99);

        band.Objects.Add(text);
        page.ReportTitle = band;
        report.Pages.Add(page);
        report.Prepare();
        return report;
    }

}
