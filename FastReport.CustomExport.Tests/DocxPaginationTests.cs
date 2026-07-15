using FastReport.Export.Custom;
using System.IO;
using Xunit;

namespace FastReport.CustomExport.Tests;

public class DocxPaginationTests
{
    [Fact]
    public void AdvancedMatrixShouldNotSpillIntoExtraWordPages()
    {
        string path = ReportFixtureCatalog.GetPath("TextReport.frx");
        using Report report = new Report();
        report.Load(path);
        report.Prepare();

        using MemoryStream stream = new MemoryStream();
        using DocxExport export = new DocxExport();
        report.Export(export, stream);

        int expectedPages = report.PreparedPages.Count;
        int actualPages = WordDocxPageCounter.CountPages(stream);

        Assert.Equal(expectedPages, actualPages);
    }
}
