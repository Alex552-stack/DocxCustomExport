using FastReport.Export.Custom;
using System.Linq;
using Xunit;

namespace FastReport.CustomExport.Tests;

public class DocxLayoutGridTests
{
    [Fact]
    public void GridBuilderShouldCreateEdgesFromBandsAndTextBounds()
    {
        using Report report = CreateStructuredReport();
        using ReportPage page = report.PreparedPages.GetPage(0);

        DocxLayoutSnapshot snapshot = DocxLayoutSnapshotBuilder.Build(page);
        DocxLayoutGrid grid = DocxLayoutGridBuilder.Build(snapshot);

        Assert.Equal(0, grid.XEdges.First());
        Assert.Equal(snapshot.PageWidth, grid.XEdges.Last());
        Assert.Equal(0, grid.YEdges.First());
        Assert.Equal(snapshot.PageHeight, grid.YEdges.Last());
        Assert.Contains(grid.XEdges, x => x == 20);
        Assert.Contains(grid.XEdges, x => x == 200);
        Assert.Contains(grid.YEdges, y => y == 12);
        Assert.Contains(grid.YEdges, y => y == 32);
        Assert.True(grid.ColumnCount > 0);
        Assert.True(grid.RowCount > 0);
    }

    [Fact]
    public void GridBuilderShouldMapObjectsToStableCellSpans()
    {
        using Report report = CreateStructuredReport();
        using ReportPage page = report.PreparedPages.GetPage(0);

        DocxLayoutGrid grid = DocxLayoutGridBuilder.Build(DocxLayoutSnapshotBuilder.Build(page));
        DocxLayoutGridObject titleBand = grid.Objects.Single(x => x.Name == "ReportTitle1");
        DocxLayoutGridObject titleText = grid.Objects.Single(x => x.Name == "TitleText");
        DocxLayoutGridObject footerText = grid.Objects.Single(x => x.Name == "FooterText");

        Assert.Equal("ReportTitle1", titleText.ParentBandName);
        Assert.Equal(0, titleBand.Column);
        Assert.True(titleBand.ColumnSpan >= titleText.ColumnSpan);
        Assert.True(titleBand.RowSpan >= titleText.RowSpan);
        Assert.True(footerText.Row > titleText.Row);
        Assert.True(footerText.ColumnSpan >= 1);
    }

    [Fact]
    public void GridBuilderShouldMergeNearDuplicateEdgesUsingTolerance()
    {
        DocxLayoutSnapshot snapshot = new()
        {
            PageWidth = 500,
            PageHeight = 300,
            Bands = new[]
            {
                new DocxBandSnapshot
                {
                    Name = "Band1",
                    BandType = nameof(ReportTitleBand),
                    Left = 0,
                    Top = 0,
                    Width = 500,
                    Height = 50
                }
            },
            Texts = new[]
            {
                new DocxTextSnapshot
                {
                    Name = "NearLeft",
                    Left = 20,
                    Top = 10,
                    Width = 100,
                    Height = 10
                },
                new DocxTextSnapshot
                {
                    Name = "NearLeft2",
                    Left = 20.3f,
                    Top = 30,
                    Width = 100,
                    Height = 10
                }
            }
        };

        DocxLayoutGrid grid = DocxLayoutGridBuilder.Build(snapshot, tolerance: 0.5f);

        Assert.Single(grid.XEdges.Where(x => x >= 19.5f && x <= 20.5f));
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
