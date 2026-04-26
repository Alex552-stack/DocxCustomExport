using FastReport.Export;
using System.Collections.Generic;
using System.IO;

namespace FastReport.Export.Custom;

/// <summary>
/// Minimal custom export scaffold that creates empty output files per prepared page.
/// </summary>
public class EmptyCustomExport : ExportBase
{
    private int pageIndex;
    private bool saveStreams;

    /// <summary>
    /// When enabled, generated files are returned through <see cref="ExportBase.GeneratedStreams"/>.
    /// </summary>
    public bool SaveStreams
    {
        get => saveStreams;
        set => saveStreams = value;
    }

    /// <inheritdoc />
    protected override string GetFileFilter()
    {
        return "Custom empty export (*.fce)|*.fce";
    }

    /// <inheritdoc />
    protected override void Start()
    {
        base.Start();
        pageIndex = 0;
        GeneratedStreams = new List<Stream>();
    }

    /// <inheritdoc />
    protected override void ExportPageBegin(ReportPage page)
    {
        base.ExportPageBegin(page);
        pageIndex++;

        string outputFileName = GetOutputFileName(pageIndex);
        if (saveStreams)
        {
            GeneratedFiles.Add(outputFileName);
            GeneratedStreams.Add(new MemoryStream());
            return;
        }

        string? directory = Path.GetDirectoryName(outputFileName);
        if (!string.IsNullOrEmpty(directory))
            Directory.CreateDirectory(directory);

        using FileStream fileStream = new FileStream(outputFileName, FileMode.Create, FileAccess.Write);
        GeneratedFiles.Add(outputFileName);
    }

    private string GetOutputFileName(int currentPage)
    {
        string baseName = string.IsNullOrWhiteSpace(FileName)
            ? "custom-export"
            : Path.Combine(
                Path.GetDirectoryName(FileName) ?? string.Empty,
                Path.GetFileNameWithoutExtension(FileName));

        return baseName + $".page{currentPage}.fce";
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="EmptyCustomExport"/> class.
    /// </summary>
    public EmptyCustomExport()
    {
        HasMultipleFiles = true;
        saveStreams = true;
    }
}
