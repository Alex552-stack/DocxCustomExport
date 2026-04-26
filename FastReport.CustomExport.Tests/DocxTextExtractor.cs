using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FastReport.CustomExport.Tests;

internal static class DocxTextExtractor
{
    public static IReadOnlyList<string> ExtractAllText(Stream docxStream)
    {
        if (docxStream.CanSeek)
            docxStream.Position = 0;

        using WordprocessingDocument document = WordprocessingDocument.Open(docxStream, false);
        IEnumerable<string> texts = document.MainDocumentPart?
            .Document?
            .Descendants<Text>()
            .Select(text => TextNormalization.NormalizeText(text.Text))
            .Where(text => !string.IsNullOrWhiteSpace(text))
            ?? Enumerable.Empty<string>();

        return texts.ToList();
    }

    public static string ExtractAsSingleText(Stream docxStream)
    {
        return string.Join("\n", ExtractAllText(docxStream));
    }
}
