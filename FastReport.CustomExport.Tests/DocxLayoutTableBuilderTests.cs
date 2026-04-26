using DocumentFormat.OpenXml.Wordprocessing;
using FastReport.Export.Custom;
using System.Linq;
using Xunit;
using WordTable = DocumentFormat.OpenXml.Wordprocessing.Table;

namespace FastReport.CustomExport.Tests;

public class DocxLayoutTableBuilderTests
{
    [Fact]
    public void TableBuilderShouldCreateFixedGridMatchingLayoutEdges()
    {
        using Report report = CreateStructuredReport();
        using ReportPage page = report.PreparedPages.GetPage(0);
        DocxLayoutSnapshot snapshot = DocxLayoutSnapshotBuilder.Build(page);
        DocxLayoutGrid grid = DocxLayoutGridBuilder.Build(snapshot);

        WordTable table = DocxLayoutTableBuilder.Build(grid, snapshot);

        TableGrid gridDefinition = table.GetFirstChild<TableGrid>()!;
        Assert.NotNull(grid.GetType());
        Assert.Equal(grid.ColumnCount, gridDefinition.Elements<GridColumn>().Count());
        Assert.Equal(grid.RowCount, table.Elements<TableRow>().Count());
        Assert.Equal(TableLayoutValues.Fixed, table.GetFirstChild<TableProperties>()!.GetFirstChild<TableLayout>()!.Type!.Value);
    }

    [Fact]
    public void TableBuilderShouldPlaceTextIntoOriginCells()
    {
        using Report report = CreateStructuredReport();
        using ReportPage page = report.PreparedPages.GetPage(0);
        DocxLayoutSnapshot snapshot = DocxLayoutSnapshotBuilder.Build(page);
        DocxLayoutGrid grid = DocxLayoutGridBuilder.Build(snapshot);

        WordTable table = DocxLayoutTableBuilder.Build(grid, snapshot);
        string[] cellTexts = table.Descendants<TableCell>()
            .Select(cell => string.Concat(cell.Descendants<Text>().Select(x => x.Text)))
            .Where(text => !string.IsNullOrWhiteSpace(text))
            .ToArray();

        Assert.Contains("Title", cellTexts);
        Assert.Contains("Body", cellTexts);
        Assert.Contains("Footer", cellTexts);
    }

    [Fact]
    public void TableBuilderShouldEmitVerticalMergeForMultiRowObjects()
    {
        DocxLayoutSnapshot snapshot = new()
        {
            PageWidth = 300,
            PageHeight = 200,
            Texts = new[]
            {
                new DocxTextSnapshot
                {
                    Name = "Tall",
                    Text = "Tall",
                    Left = 10,
                    Top = 10,
                    Width = 80,
                    Height = 60
                },
                new DocxTextSnapshot
                {
                    Name = "Bottom",
                    Text = "Bottom",
                    Left = 100,
                    Top = 40,
                    Width = 50,
                    Height = 20
                },
                new DocxTextSnapshot
                {
                    Name = "Later",
                    Text = "Later",
                    Left = 100,
                    Top = 100,
                    Width = 50,
                    Height = 20
                }
            }
        };

        DocxLayoutGrid grid = DocxLayoutGridBuilder.Build(snapshot);
        WordTable table = DocxLayoutTableBuilder.Build(grid, snapshot);

        Assert.Contains(table.Descendants<VerticalMerge>(), x => x.Val?.Value == MergedCellValues.Restart);
        Assert.Contains(table.Descendants<VerticalMerge>(), x => x.Val == null);
    }

    private static Report CreateStructuredReport()
    {
        Report report = new Report();
        ReportPage page = new ReportPage
        {
            PaperWidth = 210,
            PaperHeight = 297
        };

        ReportTitleBand titleBand = new ReportTitleBand
        {
            Name = "ReportTitle1",
            Height = 70
        };
        titleBand.Objects.Add(new TextObject
        {
            Name = "TitleText",
            Text = "Title",
            Left = 20,
            Top = 12,
            Width = 180,
            Height = 20
        });

        DataBand dataBand = new DataBand
        {
            Name = "Data1",
            Height = 120
        };
        dataBand.Objects.Add(new TextObject
        {
            Name = "BodyText",
            Text = "Body",
            Left = 30,
            Top = 10,
            Width = 160,
            Height = 20
        });

        PageFooterBand footerBand = new PageFooterBand
        {
            Name = "PageFooter1",
            Height = 28
        };
        footerBand.Objects.Add(new TextObject
        {
            Name = "FooterText",
            Text = "Footer",
            Left = 10,
            Top = 0,
            Width = 100,
            Height = 20
        });

        page.ReportTitle = titleBand;
        page.Bands.Add(dataBand);
        page.PageFooter = footerBand;
        report.Pages.Add(page);
        report.Prepare();
        return report;
    }
}
