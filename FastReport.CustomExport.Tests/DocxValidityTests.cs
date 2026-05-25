using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FastReport.Export.Custom;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;

namespace FastReport.CustomExport.Tests;

public class DocxValidityTests
{
    private const int MaxValidationErrorsToReport = 20;

    [Theory]
    [MemberData(nameof(LocalPreparedFixtures))]
    public void LocalPreparedReportShouldExportValidDocxPackage(PreparedReportFixtureInfo fixture)
    {
        using Report report = PreparedReportLoader.LoadPreparedReport(fixture);
        using MemoryStream docxStream = ExportToDocx(report);

        using WordprocessingDocument document = WordprocessingDocument.Open(docxStream, false);

        MainDocumentPart mainDocumentPart = document.MainDocumentPart
            ?? throw new InvalidDataException($"DOCX package has no main document part for fixture '{fixture.Name}'.");
        Assert.NotNull(mainDocumentPart.Document);

        OpenXmlValidator validator = new OpenXmlValidator();
        string[] validationErrors = validator.Validate(document)
            .Take(MaxValidationErrorsToReport)
            .Select(error => $"{error.Path?.XPath}: {error.Description}")
            .ToArray();

        Assert.True(
            validationErrors.Length == 0,
            $"DOCX package is not OpenXML-valid for fixture '{fixture.Name}'. Showing first {MaxValidationErrorsToReport} errors:{System.Environment.NewLine}{string.Join(System.Environment.NewLine, validationErrors)}");
    }

    public static IEnumerable<object[]> LocalPreparedFixtures()
    {
        return PreparedReportFixtureCatalog.GetLocalPreparedFixturesAsTheoryData();
    }

    private static MemoryStream ExportToDocx(Report report)
    {
        MemoryStream stream = new MemoryStream();
        using DocxExport export = new DocxExport();
        report.Export(export, stream);
        stream.Position = 0;
        return stream;
    }
}
