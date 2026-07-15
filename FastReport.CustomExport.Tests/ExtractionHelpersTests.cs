using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using Xunit;

namespace FastReport.CustomExport.Tests;

public class ExtractionHelpersTests
{
    [Fact]
    public void PreparedReportTextExtractorReturnsKnownText()
    {
        PreparedReportFixtureInfo fixture = new PreparedReportFixtureInfo
        {
            Kind = ReportFixtureKind.Frx,
            Name = "TextReport.frx",
            Path = ReportFixtureCatalog.GetPath("TextReport.frx")
        };

        using Report report = PreparedReportLoader.LoadPreparedReport(fixture);

        string extractedText = PreparedReportTextExtractor.ExtractAsSingleText(report);

        Assert.Contains("TextObject", extractedText);
        Assert.Contains("TEXT, BORDER, FILL", extractedText);
    }

    [Fact]
    public void DocxTextExtractorReturnsDocumentText()
    {
        using MemoryStream stream = CreateDocx("Alpha", "Beta", "Gamma");

        string extractedText = DocxTextExtractor.ExtractAsSingleText(stream);

        Assert.Contains("Alpha", extractedText);
        Assert.Contains("Beta", extractedText);
        Assert.Contains("Gamma", extractedText);
    }

    private static MemoryStream CreateDocx(params string[] lines)
    {
        MemoryStream stream = new MemoryStream();

        using (WordprocessingDocument document = WordprocessingDocument.Create(
            stream,
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document,
            true))
        {
            MainDocumentPart mainDocumentPart = document.AddMainDocumentPart();
            Body body = new Body();

            foreach (string line in lines)
            {
                body.AppendChild(new Paragraph(
                    new Run(
                        new Text(line))));
            }

            mainDocumentPart.Document = new Document(body);
            mainDocumentPart.Document.Save();
        }

        stream.Position = 0;
        return stream;
    }
}
