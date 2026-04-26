using FastReport.Export.Custom;
using System.IO;
using Xunit;

namespace FastReport.CustomExport.Tests;

public class DocxMarginTests
{
    [Fact]
    public void ExportedDocxShouldPreservePageMargins()
    {
        using Report report = CreateSinglePageReportWithMargins(
            left: 11,
            top: 17,
            right: 23,
            bottom: 29);

        using MemoryStream stream = new MemoryStream();
        DocxExport export = new DocxExport();
        report.Export(export, stream);

        DocxMargins margins = DocxLayoutReader.ReadMargins(stream);

        Assert.Equal(624u, margins.Left);
        Assert.Equal(964u, margins.Top);
        Assert.Equal(1304u, margins.Right);
        Assert.Equal(1644u, margins.Bottom);
    }

    private static Report CreateSinglePageReportWithMargins(float left, float top, float right, float bottom)
    {
        Report report = new Report();
        ReportPage page = new ReportPage
        {
            LeftMargin = left,
            TopMargin = top,
            RightMargin = right,
            BottomMargin = bottom
        };

        ReportTitleBand band = new ReportTitleBand
        {
            Name = "ReportTitle1",
            Height = 40
        };

        TextObject text = new TextObject
        {
            Name = "Text1",
            Text = "Margin probe",
            Width = 300,
            Height = 20
        };

        band.Objects.Add(text);
        page.ReportTitle = band;
        report.Pages.Add(page);
        report.Prepare();
        return report;
    }
}
