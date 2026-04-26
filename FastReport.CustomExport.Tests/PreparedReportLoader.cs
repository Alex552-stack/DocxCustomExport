namespace FastReport.CustomExport.Tests;

internal static class PreparedReportLoader
{
    public static Report LoadPreparedReport(PreparedReportFixtureInfo fixture)
    {
        Report report = new Report();

        switch (fixture.Kind)
        {
            case ReportFixtureKind.Frx:
                report.Load(fixture.Path);
                report.Prepare();
                break;

            case ReportFixtureKind.Fpx:
                report.LoadPrepared(fixture.Path);
                break;
        }

        return report;
    }
}
