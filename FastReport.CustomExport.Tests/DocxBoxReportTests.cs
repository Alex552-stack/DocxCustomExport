using FastReport.Export.Custom;
using System.IO;
using System.Linq;
using Xunit;

namespace FastReport.CustomExport.Tests;

public class DocxBoxReportTests
{
    [Fact]
    public void BoxReportShouldEmitDrawableComponentsAsImages()
    {
        string path = ReportFixtureCatalog.GetPath("Reports", "Box.frx");
        using Report report = new Report();
        report.Load(path);
        report.Prepare();

        using MemoryStream stream = new MemoryStream();
        using DocxExport export = new DocxExport();
        report.Export(export, stream);

        string[] bodyTexts = DocxLayoutReader.ReadParagraphLayouts(stream)
            .Select(x => x.Text)
            .ToArray();

        Assert.Contains("Box Package with Pharmacode and EAN", bodyTexts);
        Assert.True(DocxLayoutReader.ReadBodyImageCount(stream) >= 5);
    }
}
