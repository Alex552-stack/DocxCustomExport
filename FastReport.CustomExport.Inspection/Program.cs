using FastReport.Export.Custom;
using System.Collections.Concurrent;
using System.Diagnostics;

namespace FastReport.CustomExport.Inspection;

internal static class Program
{
    public static async Task<int> Main(string[] args)
    {
        InspectionOptions options = InspectionOptions.Parse(args);
        string testReportsRoot = Path.Combine(AppContext.BaseDirectory, "TestReports");
        string outputRoot = options.ResolveOutputRoot();

        if (!Directory.Exists(testReportsRoot))
        {
            Console.Error.WriteLine($"Test reports folder not found: {testReportsRoot}");
            return 1;
        }

        List<InspectionReportInput> reports = GetReportInputs(testReportsRoot);

        if (!string.IsNullOrWhiteSpace(options.Filter))
        {
            reports = reports
                .Where(report => report.RelativeName.Contains(options.Filter, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        if (options.Limit.HasValue)
            reports = reports.Take(options.Limit.Value).ToList();

        if (reports.Count == 0)
        {
            Console.WriteLine("No FRX or FPX reports matched the requested filters.");
            return 0;
        }

        Directory.CreateDirectory(outputRoot);
        Console.WriteLine($"Exporting {reports.Count} FRX/FPX reports to: {outputRoot}");
        Console.WriteLine($"Parallelism: {options.Parallelism}");

        ConcurrentBag<InspectionExportResult> results = new ConcurrentBag<InspectionExportResult>();
        ParallelOptions parallelOptions = new ParallelOptions
        {
            MaxDegreeOfParallelism = options.Parallelism
        };

        await Parallel.ForEachAsync(reports, parallelOptions, (report, cancellationToken) =>
        {
            InspectionExportResult result = ExportPair(report, outputRoot);
            results.Add(result);

            string status = result.Success ? "OK" : "FAIL";
            Console.WriteLine($"[{status}] {result.RelativeName}");
            return ValueTask.CompletedTask;
        });

        List<InspectionExportResult> orderedResults = results
            .OrderBy(result => result.RelativeName, StringComparer.OrdinalIgnoreCase)
            .ToList();

        string summaryPath = Path.Combine(outputRoot, "_summary.csv");
        WriteSummary(summaryPath, orderedResults);

        int successCount = orderedResults.Count(result => result.Success);
        int failureCount = orderedResults.Count - successCount;

        Console.WriteLine($"Finished. Success: {successCount}. Failed: {failureCount}.");
        Console.WriteLine($"Summary: {summaryPath}");

        return failureCount == 0 ? 0 : 2;
    }

    private static List<InspectionReportInput> GetReportInputs(string testReportsRoot)
    {
        List<InspectionReportInput> reports = new List<InspectionReportInput>();
        string frxRoot = Path.Combine(testReportsRoot, "Reports");

        if (Directory.Exists(frxRoot))
        {
            reports.AddRange(Directory.GetFiles(frxRoot, "*.frx", SearchOption.AllDirectories)
                .Select(path => InspectionReportInput.Create(path, frxRoot, InspectionReportKind.Frx)));
        }

        reports.AddRange(Directory.GetFiles(testReportsRoot, "*.fpx", SearchOption.AllDirectories)
            .Select(path => InspectionReportInput.Create(path, testReportsRoot, InspectionReportKind.Fpx)));

        return reports
            .OrderBy(report => report.RelativeName, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static InspectionExportResult ExportPair(InspectionReportInput reportInput, string outputRoot)
    {
        string relativeDirectory = Path.GetDirectoryName(reportInput.RelativeName) ?? string.Empty;
        string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(reportInput.RelativeName);

        Stopwatch stopwatch = Stopwatch.StartNew();

        try
        {
            string pairDirectory = GetPairDirectory(outputRoot, "successful", relativeDirectory, fileNameWithoutExtension);
            string sourcePath = Path.Combine(pairDirectory, $"{fileNameWithoutExtension}.source{reportInput.Extension}");
            string fpxPath = Path.Combine(pairDirectory, $"{fileNameWithoutExtension}.fpx");
            string docxPath = Path.Combine(pairDirectory, $"{fileNameWithoutExtension}.docx");

            Directory.CreateDirectory(pairDirectory);
            File.Copy(reportInput.Path, sourcePath, true);

            using Report report = new Report();
            if (reportInput.Kind == InspectionReportKind.Frx)
            {
                report.Load(reportInput.Path);
                report.Prepare();
                report.SavePrepared(fpxPath);
            }
            else
            {
                report.LoadPrepared(reportInput.Path);
                File.Copy(reportInput.Path, fpxPath, true);
            }

            using FileStream docxStream = File.Create(docxPath);
            using DocxExport export = new DocxExport();
            report.Export(export, docxStream);

            stopwatch.Stop();

            return new InspectionExportResult
            {
                RelativeName = reportInput.RelativeName,
                SourcePath = sourcePath,
                FpxPath = fpxPath,
                DocxPath = docxPath,
                Success = true,
                DurationMs = stopwatch.ElapsedMilliseconds
            };
        }
        catch (Exception ex)
        {
            stopwatch.Stop();

            string successfulPairDirectory = GetPairDirectory(outputRoot, "successful", relativeDirectory, fileNameWithoutExtension);
            string successfulFpxPath = Path.Combine(successfulPairDirectory, $"{fileNameWithoutExtension}.fpx");
            string pairDirectory = GetPairDirectory(outputRoot, "failed", relativeDirectory, fileNameWithoutExtension);
            string sourcePath = Path.Combine(pairDirectory, $"{fileNameWithoutExtension}.source{reportInput.Extension}");
            string fpxPath = Path.Combine(pairDirectory, $"{fileNameWithoutExtension}.fpx");
            string docxPath = Path.Combine(pairDirectory, $"{fileNameWithoutExtension}.docx");

            Directory.CreateDirectory(pairDirectory);
            File.Copy(reportInput.Path, sourcePath, true);
            if (File.Exists(successfulFpxPath))
                File.Copy(successfulFpxPath, fpxPath, true);
            else if (reportInput.Kind == InspectionReportKind.Fpx)
                File.Copy(reportInput.Path, fpxPath, true);

            DeleteDirectoryIfExists(successfulPairDirectory);

            return new InspectionExportResult
            {
                RelativeName = reportInput.RelativeName,
                SourcePath = sourcePath,
                FpxPath = fpxPath,
                DocxPath = docxPath,
                Success = false,
                DurationMs = stopwatch.ElapsedMilliseconds,
                Error = ex.ToString()
            };
        }
    }

    private static void DeleteDirectoryIfExists(string path)
    {
        if (Directory.Exists(path))
            Directory.Delete(path, true);
    }

    private static string GetPairDirectory(string outputRoot, string bucket, string relativeDirectory, string fileNameWithoutExtension)
    {
        return Path.Combine(outputRoot, bucket, relativeDirectory, fileNameWithoutExtension);
    }

    private static void WriteSummary(string summaryPath, IReadOnlyList<InspectionExportResult> results)
    {
        using StreamWriter writer = new StreamWriter(summaryPath, false);
        writer.WriteLine("RelativeName,Success,DurationMs,SourcePath,FpxPath,DocxPath,Error");

        foreach (InspectionExportResult result in results)
        {
            writer.WriteLine(string.Join(",",
                Csv(result.RelativeName),
                result.Success ? "true" : "false",
                result.DurationMs.ToString(),
                Csv(result.SourcePath),
                Csv(result.FpxPath),
                Csv(result.DocxPath),
                Csv(result.Error ?? string.Empty)));
        }
    }

    private static string Csv(string value)
    {
        string escaped = value.Replace("\"", "\"\"");
        return $"\"{escaped}\"";
    }
}

internal enum InspectionReportKind
{
    Frx,
    Fpx
}

internal sealed class InspectionReportInput
{
    public string Path { get; init; } = string.Empty;

    public string RelativeName { get; init; } = string.Empty;

    public string Extension { get; init; } = string.Empty;

    public InspectionReportKind Kind { get; init; }

    public static InspectionReportInput Create(string path, string root, InspectionReportKind kind)
    {
        return new InspectionReportInput
        {
            Path = path,
            RelativeName = System.IO.Path.GetRelativePath(root, path),
            Extension = System.IO.Path.GetExtension(path),
            Kind = kind
        };
    }
}

internal sealed class InspectionOptions
{
    public string? OutputRoot { get; set; }

    public string? Filter { get; set; }

    public int? Limit { get; set; }

    public int Parallelism { get; set; } = Math.Max(1, Environment.ProcessorCount / 2);

    public string ResolveOutputRoot()
    {
        if (!string.IsNullOrWhiteSpace(OutputRoot))
            return Path.GetFullPath(OutputRoot);

        string timestamp = DateTime.Now.ToString("yyyyMMdd-HHmmss");
        return Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..",
            "..",
            "..",
            "..",
            "ManualInspectionOutput",
            timestamp));
    }

    public static InspectionOptions Parse(string[] args)
    {
        InspectionOptions options = new InspectionOptions();

        for (int i = 0; i < args.Length; i++)
        {
            string arg = args[i];
            switch (arg)
            {
                case "--output":
                    options.OutputRoot = ReadValue(args, ref i, arg);
                    break;

                case "--filter":
                    options.Filter = ReadValue(args, ref i, arg);
                    break;

                case "--limit":
                    options.Limit = Int32.Parse(ReadValue(args, ref i, arg));
                    break;

                case "--parallelism":
                    options.Parallelism = Math.Max(1, Int32.Parse(ReadValue(args, ref i, arg)));
                    break;

                case "--help":
                case "-h":
                    PrintUsage();
                    Environment.Exit(0);
                    break;

                default:
                    throw new ArgumentException($"Unknown argument: {arg}");
            }
        }

        return options;
    }

    private static string ReadValue(string[] args, ref int index, string optionName)
    {
        if (index + 1 >= args.Length)
            throw new ArgumentException($"Missing value for {optionName}");

        index++;
        return args[index];
    }

    private static void PrintUsage()
    {
        Console.WriteLine("FastReport.CustomExport.Inspection");
        Console.WriteLine("  --output <path>       Output folder root");
        Console.WriteLine("  --filter <text>       Export only reports whose relative path contains the text");
        Console.WriteLine("  --limit <n>           Export only the first n matching reports");
        Console.WriteLine("  --parallelism <n>     Max number of reports exported in parallel");
    }
}

internal sealed class InspectionExportResult
{
    public string RelativeName { get; init; } = string.Empty;

    public string SourcePath { get; init; } = string.Empty;

    public string FpxPath { get; init; } = string.Empty;

    public string DocxPath { get; init; } = string.Empty;

    public bool Success { get; init; }

    public long DurationMs { get; init; }

    public string? Error { get; init; }
}
