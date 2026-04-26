using System.Collections.Generic;
using System.Linq;

namespace FastReport.CustomExport.Tests;

internal static class PreparedReportTextExtractor
{
    public static IReadOnlyList<string> ExtractAllText(Report report)
    {
        List<string> texts = new List<string>();

        for (int pageIndex = 0; pageIndex < report.PreparedPages.Count; pageIndex++)
        {
            using ReportPage page = report.PreparedPages.GetPage(pageIndex);
            foreach (Base obj in page.AllObjects)
            {
                if (obj is TextObjectBase textObject)
                {
                    string normalized = TextNormalization.NormalizeText(textObject.Text);
                    if (!string.IsNullOrWhiteSpace(normalized))
                        texts.Add(normalized);
                }
            }
        }

        return texts;
    }

    public static string ExtractAsSingleText(Report report)
    {
        return string.Join("\n", ExtractAllText(report));
    }

    public static bool ContainsAllSourceText(Report report, string targetText)
    {
        string normalizedTarget = TextNormalization.NormalizeText(targetText);
        return ExtractAllText(report).All(text => normalizedTarget.Contains(text));
    }
}
