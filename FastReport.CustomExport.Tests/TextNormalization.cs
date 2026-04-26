using System.Text.RegularExpressions;

namespace FastReport.CustomExport.Tests;

internal static class TextNormalization
{
    private static readonly Regex WhitespaceRegex = new Regex(@"\s+", RegexOptions.Compiled);

    public static string NormalizeText(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return string.Empty;

        return WhitespaceRegex.Replace(text, " ").Trim();
    }
}
