using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FastReport.CustomExport.Tests;

internal static class PreparedReportFixtureCatalog
{
    public static IReadOnlyList<PreparedReportFixtureInfo> GetAllFixtures()
    {
        List<PreparedReportFixtureInfo> fixtures = new List<PreparedReportFixtureInfo>();

        fixtures.AddRange(
            ReportFixtureCatalog.GetFrxReports().Select(path => new PreparedReportFixtureInfo
            {
                Kind = ReportFixtureKind.Frx,
                Name = System.IO.Path.GetFileName(path),
                Path = path
            }));

        fixtures.AddRange(
            ReportFixtureCatalog.GetPreparedReports().Select(path => new PreparedReportFixtureInfo
            {
                Kind = ReportFixtureKind.Fpx,
                Name = System.IO.Path.GetFileName(path),
                Path = path
            }));

        return fixtures;
    }

    public static IEnumerable<object[]> GetAllFixturesAsTheoryData()
    {
        return GetAllFixtures().Select(fixture => new object[] { fixture });
    }

    public static IEnumerable<object[]> GetPreparedFixturesAsTheoryData()
    {
        return GetAllFixtures()
            .Where(fixture => fixture.Kind == ReportFixtureKind.Fpx)
            .Select(fixture => new object[] { fixture });
    }

    public static IEnumerable<object[]> GetLocalPreparedFixturesAsTheoryData()
    {
        return ReportFixtureCatalog.GetLocalPreparedReports()
            .Select(path => new PreparedReportFixtureInfo
            {
                Kind = ReportFixtureKind.Fpx,
                Name = System.IO.Path.GetFileName(path),
                Path = path
            })
            .Select(fixture => new object[] { fixture });
    }
}
