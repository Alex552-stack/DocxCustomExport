using FastReport.Export.Custom;
using System.IO;
using Xunit;

namespace FastReport.CustomExport.Tests;

public class DocxExportContractTests
{
    private const string PendingPaginationReason =
        "Pending DOCX export implementation and renderer-backed page count verification. Pagination is non-blocking for now.";

    [Theory(Skip = PendingPaginationReason)]
    [MemberData(nameof(PreparedFixtures))]
    public void ExportedDocxShouldExposeComparableRenderedPageCount(PreparedReportFixtureInfo fixture)
    {
        using Report report = PreparedReportLoader.LoadPreparedReport(fixture);
        using MemoryStream docxStream = ExportToDocx(report);

        int sourcePageCount = report.PreparedPages.Count;
        int docxRenderedPageCount = GetRenderedDocxPageCount(docxStream);

        Assert.Equal(sourcePageCount, docxRenderedPageCount);
    }

    [Theory]
    [MemberData(nameof(PreparedFixtures))]
    public void ExportedDocxShouldContainAllPreparedText(PreparedReportFixtureInfo fixture)
    {
        using Report report = PreparedReportLoader.LoadPreparedReport(fixture);
        using MemoryStream docxStream = ExportToDocx(report);

        string docxText = DocxTextExtractor.ExtractAsSingleText(docxStream);

        Assert.True(
            PreparedReportTextExtractor.ContainsAllSourceText(report, docxText),
            $"DOCX text is missing source text for fixture '{fixture.Name}'.");
    }

    public static IEnumerable<object[]> PreparedFixtures()
    {
        return PreparedReportFixtureCatalog.GetPreparedFixturesAsTheoryData();
    }

    private static MemoryStream ExportToDocx(Report report)
    {
        MemoryStream stream = new MemoryStream();

        // The future implementation should follow the same export architecture style
        // used by HTMLExport and PDFSimpleExport rather than introducing a parallel pipeline.
        DocxExport export = new DocxExport();
        report.Export(export, stream);
        stream.Position = 0;
        return stream;
    }

    private static int GetRenderedDocxPageCount(Stream docxStream)
    {
        throw new System.NotImplementedException(
            "Page count verification will use a renderer-backed DOCX pipeline, e.g. headless LibreOffice.");
    }
}
