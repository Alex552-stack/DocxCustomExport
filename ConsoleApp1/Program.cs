using FastReport;
using FastReport.Export.Custom;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;

string fixture = args.Length > 0
    ? args[0]
    : @"C:\Users\alex\Desktop\Fr\FastReport\Tools\FastReport.CustomExport.Tests\bin\Debug\net6.0\TestReports\Prepared\Avalonia\Simple List.fpx";

using Report report = new Report();
report.LoadPrepared(fixture);

using MemoryStream stream = new MemoryStream();
using DocxExport export = new DocxExport();
report.Export(export, stream);
stream.Position = 0;

using WordprocessingDocument document = WordprocessingDocument.Open(stream, false);
List<string> texts = document.MainDocumentPart?.Document?
    .Descendants<Text>()
    .Select(x => Normalize(x.Text))
    .Where(x => !string.IsNullOrWhiteSpace(x))
    .ToList() ?? new List<string>();

Console.WriteLine($"Texts: {texts.Count}");
foreach (string text in texts.Take(50))
    Console.WriteLine(text);

string allDocxText = Normalize(string.Join("\n", texts));
List<string> sourceTexts = new List<string>();
for (int pageIndex = 0; pageIndex < report.PreparedPages.Count; pageIndex++)
{
    using ReportPage page = report.PreparedPages.GetPage(pageIndex);
    foreach (Base obj in page.AllObjects)
    {
        if (obj is TextObjectBase textObject)
        {
            string normalized = Normalize(textObject.Text);
            if (!string.IsNullOrWhiteSpace(normalized))
                sourceTexts.Add(normalized);
        }
    }
}

Console.WriteLine("Missing:");
foreach (string missing in sourceTexts.Where(x => !allDocxText.Contains(x)).Distinct().Take(50))
    Console.WriteLine(missing);

static string Normalize(string? text)
{
    if (string.IsNullOrWhiteSpace(text))
        return string.Empty;

    return Regex.Replace(text, "\\s+", " ").Trim();
}
