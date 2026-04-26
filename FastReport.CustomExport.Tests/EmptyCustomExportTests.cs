using FastReport.Export.Custom;
using System.IO;
using Xunit;

namespace FastReport.CustomExport.Tests;

public class EmptyCustomExportTests
{
    [Fact]
    public void ExportFromFrxReturnsGeneratedStreams()
    {
        using Report report = new Report();
        report.Load(ReportFixtureCatalog.GetPath("Reports", "Text.frx"));
        report.Prepare();

        EmptyCustomExport export = new EmptyCustomExport
        {
            SaveStreams = true
        };

        using MemoryStream output = new MemoryStream();
        report.Export(export, output);

        Assert.NotEmpty(export.GeneratedFiles);
        Assert.Equal(export.GeneratedFiles.Count, export.GeneratedStreams.Count);
        Assert.All(export.GeneratedStreams, stream => Assert.NotNull(stream));
    }

    [Fact]
    public void ExportFromPreparedReportReturnsGeneratedStreams()
    {
        using Report report = new Report();
        report.LoadPrepared(ReportFixtureCatalog.GetPath("Prepared", "Avalonia", "Simple List.fpx"));

        EmptyCustomExport export = new EmptyCustomExport
        {
            SaveStreams = true
        };

        using MemoryStream output = new MemoryStream();
        report.Export(export, output);

        Assert.NotEmpty(export.GeneratedFiles);
        Assert.Equal(export.GeneratedFiles.Count, report.PreparedPages.Count);
        Assert.All(export.GeneratedFiles, file => Assert.EndsWith(".fce", file));
    }

    [Fact]
    public void FullFixtureCatalogIsAvailable()
    {
        IReadOnlyList<string> frxReports = ReportFixtureCatalog.GetFrxReports();
        IReadOnlyList<string> preparedReports = ReportFixtureCatalog.GetPreparedReports();

        Assert.True(frxReports.Count >= 100);
        Assert.True(preparedReports.Count >= 7);
    }
}
