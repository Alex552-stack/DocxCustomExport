namespace FastReport.CustomExport.Tests;

public sealed class PreparedReportFixtureInfo
{
    public string Name { get; init; } = string.Empty;

    public string Path { get; init; } = string.Empty;

    public ReportFixtureKind Kind { get; init; }

    public override string ToString()
    {
        return $"{Kind}: {Name}";
    }
}

public enum ReportFixtureKind
{
    Frx,
    Fpx
}
