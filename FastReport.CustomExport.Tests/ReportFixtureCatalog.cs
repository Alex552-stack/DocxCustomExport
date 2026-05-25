using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FastReport.CustomExport.Tests;

internal static class ReportFixtureCatalog
{
    private static string Root => Path.Combine(AppContext.BaseDirectory, "TestReports");

    public static IReadOnlyList<string> GetFrxReports()
    {
        return GetFiles(Path.Combine(Root, "Reports"), "*.frx");
    }

    public static IReadOnlyList<string> GetPreparedReports()
    {
        List<string> reports = new List<string>();
        reports.AddRange(GetFiles(Path.Combine(Root, "Prepared"), "*.fpx"));
        reports.AddRange(GetFiles(Root, "*.fpx"));
        return reports;
    }

    public static IReadOnlyList<string> GetLocalPreparedReports()
    {
        if (!Directory.Exists(Root))
            return new List<string>();

        return Directory.GetFiles(Root, "*.fpx", SearchOption.TopDirectoryOnly)
            .OrderBy(file => file)
            .ToList();
    }

    public static string GetPath(params string[] parts)
    {
        return Path.Combine(new[] { Root }.Concat(parts).ToArray());
    }

    private static List<string> GetFiles(string path, string pattern)
    {
        if (!Directory.Exists(path))
            return new List<string>();

        return Directory.GetFiles(path, pattern, SearchOption.AllDirectories)
            .OrderBy(file => file)
            .ToList();
    }
}
