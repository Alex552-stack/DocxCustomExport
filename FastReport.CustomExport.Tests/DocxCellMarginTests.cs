using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FastReport.Export.Custom;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Xunit;
using WordTable = DocumentFormat.OpenXml.Wordprocessing.Table;

namespace FastReport.CustomExport.Tests;

public class DocxCellMarginTests
{
    [Fact]
    public void SpannedTextCellShouldPreservePaddingAsCellMargins()
    {
        using Report report = CreateReportWithSpannedPaddedCell();
        using MemoryStream stream = new MemoryStream();
        using DocxExport export = new DocxExport();
        report.Export(export, stream);
        stream.Position = 0;

        using WordprocessingDocument document = WordprocessingDocument.Open(stream, false);
        TableCell cell = document.MainDocumentPart!.Document.Descendants<TableCell>()
            .First(x => string.Concat(x.Descendants<Text>().Select(text => text.Text)) == "Padded");

        TableCellProperties properties = cell.GetFirstChild<TableCellProperties>()!;
        Assert.True(properties.GetFirstChild<GridSpan>()?.Val?.Value > 1);

        TableCellMargin margin = properties.GetFirstChild<TableCellMargin>()!;
        Assert.NotNull(margin);
        Assert.Equal(PixelsToTwips(6).ToString(), margin.TopMargin!.Width!.Value);
        Assert.Equal(PixelsToTwips(10).ToString(), margin.LeftMargin!.Width!.Value);
        Assert.Equal(PixelsToTwips(8).ToString(), margin.BottomMargin!.Width!.Value);
        Assert.Equal(PixelsToTwips(12).ToString(), margin.RightMargin!.Width!.Value);
    }

    [Fact]
    public void VerticalMergeContinuationCellsShouldPreservePaddingAsCellMargins()
    {
        using Report report = CreateReportWithVerticallySpannedPaddedCell();
        using MemoryStream stream = new MemoryStream();
        using DocxExport export = new DocxExport();
        report.Export(export, stream);
        stream.Position = 0;

        using WordprocessingDocument document = WordprocessingDocument.Open(stream, false);
        WordTable table = document.MainDocumentPart!.Document.Descendants<WordTable>().First();
        TableCell originCell = table.Descendants<TableCell>()
            .First(x => string.Concat(x.Descendants<Text>().Select(text => text.Text)) == "Padded");
        TableRow originRow = originCell.Ancestors<TableRow>().First();
        int originColumn = GetColumnIndex(originRow, originCell);
        TableCell continuationCell = GetCellAtColumn(originRow.NextSibling<TableRow>()!, originColumn)!;

        Assert.Equal(MergedCellValues.Restart, originCell.GetFirstChild<TableCellProperties>()!
            .GetFirstChild<VerticalMerge>()!.Val!.Value);
        Assert.Null(continuationCell.GetFirstChild<TableCellProperties>()!
            .GetFirstChild<VerticalMerge>()!.Val);

        AssertCellMargins(continuationCell, 10, 6, 12, 8);
    }

    private static Report CreateReportWithSpannedPaddedCell()
    {
        Report report = new Report();
        ReportPage page = new ReportPage
        {
            PaperWidth = 210,
            PaperHeight = 297
        };

        ReportTitleBand band = new ReportTitleBand
        {
            Name = "ReportTitle1",
            Height = 80
        };

        band.Objects.Add(new TextObject
        {
            Name = "PaddedText",
            Text = "Padded",
            Left = 0,
            Top = 0,
            Width = 200,
            Height = 40,
            Padding = new Padding(10, 6, 12, 8)
        });

        band.Objects.Add(new TextObject
        {
            Name = "GridSplitter",
            Text = "Split",
            Left = 100,
            Top = 50,
            Width = 20,
            Height = 10
        });

        page.ReportTitle = band;
        report.Pages.Add(page);
        report.Prepare();
        return report;
    }

    private static Report CreateReportWithVerticallySpannedPaddedCell()
    {
        Report report = new Report();
        ReportPage page = new ReportPage
        {
            PaperWidth = 210,
            PaperHeight = 297
        };

        ReportTitleBand band = new ReportTitleBand
        {
            Name = "ReportTitle1",
            Height = 80
        };

        band.Objects.Add(new TextObject
        {
            Name = "PaddedText",
            Text = "Padded",
            Left = 0,
            Top = 0,
            Width = 120,
            Height = 40,
            Padding = new Padding(10, 6, 12, 8)
        });

        band.Objects.Add(new TextObject
        {
            Name = "RowSplitter",
            Text = "Split",
            Left = 140,
            Top = 20,
            Width = 20,
            Height = 10
        });

        page.ReportTitle = band;
        report.Pages.Add(page);
        report.Prepare();
        return report;
    }

    private static void AssertCellMargins(TableCell cell, int left, int top, int right, int bottom)
    {
        TableCellMargin margin = cell.GetFirstChild<TableCellProperties>()!.GetFirstChild<TableCellMargin>()!;
        Assert.NotNull(margin);
        Assert.Equal(PixelsToTwips(top).ToString(), margin.TopMargin!.Width!.Value);
        Assert.Equal(PixelsToTwips(left).ToString(), margin.LeftMargin!.Width!.Value);
        Assert.Equal(PixelsToTwips(bottom).ToString(), margin.BottomMargin!.Width!.Value);
        Assert.Equal(PixelsToTwips(right).ToString(), margin.RightMargin!.Width!.Value);
    }

    private static int GetColumnIndex(TableRow row, TableCell targetCell)
    {
        int columnIndex = 0;
        foreach (TableCell cell in row.Elements<TableCell>())
        {
            if (ReferenceEquals(cell, targetCell))
                return columnIndex;

            columnIndex += cell.GetFirstChild<TableCellProperties>()?.GetFirstChild<GridSpan>()?.Val?.Value ?? 1;
        }

        return -1;
    }

    private static TableCell? GetCellAtColumn(TableRow row, int targetColumn)
    {
        int columnIndex = 0;
        foreach (TableCell cell in row.Elements<TableCell>())
        {
            int span = cell.GetFirstChild<TableCellProperties>()?.GetFirstChild<GridSpan>()?.Val?.Value ?? 1;
            if (columnIndex == targetColumn)
                return cell;

            columnIndex += span;
        }

        return null;
    }

    private static int PixelsToTwips(float pixels)
    {
        return (int)System.Math.Round(pixels * 1440d / 96d, System.MidpointRounding.AwayFromZero);
    }
}
