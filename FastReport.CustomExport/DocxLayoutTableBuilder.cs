using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using WordTable = DocumentFormat.OpenXml.Wordprocessing.Table;

namespace FastReport.Export.Custom;

public static class DocxLayoutTableBuilder
{
    public static WordTable Build(DocxLayoutGrid grid, DocxLayoutSnapshot snapshot)
    {
        WordTable table = new();
        table.Append(BuildProperties(grid));
        table.Append(BuildGrid(grid));

        List<DocxLayoutGridObject> textObjects = grid.Objects
            .Where(x => x.Kind == nameof(TextObjectBase))
            .ToList();

        for (int rowIndex = 0; rowIndex < grid.RowCount; rowIndex++)
        {
            TableRow row = new();
            row.Append(new TableRowProperties(
                new TableRowHeight
                {
                    Val = (UInt32)ToTwips(grid.YEdges[rowIndex + 1] - grid.YEdges[rowIndex]),
                    HeightType = HeightRuleValues.Exact
                }));

            int columnIndex = 0;
            while (columnIndex < grid.ColumnCount)
            {
                List<DocxLayoutGridObject> origins = textObjects
                    .Where(x => x.Row == rowIndex && x.Column == columnIndex)
                    .OrderBy(x => x.Top)
                    .ThenBy(x => x.Left)
                    .ToList();

                if (origins.Count > 0)
                {
                    row.Append(CreateTextCell(grid, snapshot, origins));
                    columnIndex += origins.Max(x => x.ColumnSpan);
                    continue;
                }

                DocxLayoutGridObject? continuation = textObjects.FirstOrDefault(x =>
                    x.Column == columnIndex &&
                    x.Row < rowIndex &&
                    x.Row + x.RowSpan > rowIndex);

                if (continuation != null)
                {
                    row.Append(CreateContinuationCell(grid, continuation));
                    columnIndex += continuation.ColumnSpan;
                    continue;
                }

                row.Append(CreateEmptyCell(grid, columnIndex));
                columnIndex++;
            }

            table.Append(row);
        }

        return table;
    }

    private static TableProperties BuildProperties(DocxLayoutGrid grid)
    {
        return new TableProperties(
            new TableWidth
            {
                Type = TableWidthUnitValues.Dxa,
                Width = ToTwips(grid.XEdges[^1] - grid.XEdges[0]).ToString()
            },
            new TableLayout
            {
                Type = TableLayoutValues.Fixed
            },
            new TableCellMarginDefault(
                new TopMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                new TableCellLeftMargin { Width = 0, Type = TableWidthValues.Dxa },
                new BottomMargin { Width = "0", Type = TableWidthUnitValues.Dxa },
                new TableCellRightMargin { Width = 0, Type = TableWidthValues.Dxa }));
    }

    private static TableGrid BuildGrid(DocxLayoutGrid grid)
    {
        TableGrid tableGrid = new();
        for (int i = 0; i < grid.ColumnCount; i++)
        {
            tableGrid.Append(new GridColumn
            {
                Width = ToTwips(grid.XEdges[i + 1] - grid.XEdges[i]).ToString()
            });
        }

        return tableGrid;
    }

    private static TableCell CreateTextCell(
        DocxLayoutGrid grid,
        DocxLayoutSnapshot snapshot,
        IReadOnlyList<DocxLayoutGridObject> origins)
    {
        DocxLayoutGridObject origin = origins[0];
        int columnSpan = origins.Max(x => x.ColumnSpan);
        int rowSpan = origins.Max(x => x.RowSpan);
        TableCell cell = new();
        TableCellProperties properties = CreateCellProperties(grid, origin.Column, columnSpan);
        ApplyCellStyle(properties, snapshot, origin.Name);

        if (columnSpan > 1)
            properties.Append(new GridSpan { Val = columnSpan });

        if (rowSpan > 1)
            properties.Append(new VerticalMerge { Val = MergedCellValues.Restart });

        cell.Append(properties);
        foreach (DocxLayoutGridObject item in origins)
            cell.Append(CreateTextParagraph(snapshot, item.Name));
        return cell;
    }

    private static TableCell CreateContinuationCell(DocxLayoutGrid grid, DocxLayoutGridObject origin)
    {
        TableCell cell = new();
        TableCellProperties properties = CreateCellProperties(grid, origin.Column, origin.ColumnSpan);

        if (origin.ColumnSpan > 1)
            properties.Append(new GridSpan { Val = origin.ColumnSpan });

        properties.Append(new VerticalMerge());
        cell.Append(properties);
        cell.Append(new Paragraph());
        return cell;
    }

    private static TableCell CreateEmptyCell(DocxLayoutGrid grid, int columnIndex)
    {
        TableCell cell = new();
        cell.Append(CreateCellProperties(grid, columnIndex, 1));
        cell.Append(new Paragraph());
        return cell;
    }

    private static TableCellProperties CreateCellProperties(DocxLayoutGrid grid, int columnIndex, int columnSpan)
    {
        int width = ToTwips(grid.XEdges[columnIndex + columnSpan] - grid.XEdges[columnIndex]);
        return new TableCellProperties(
            new TableCellWidth
            {
                Type = TableWidthUnitValues.Dxa,
                Width = width.ToString()
            });
    }

    private static Paragraph CreateTextParagraph(DocxLayoutSnapshot snapshot, string objectName)
    {
        DocxTextSnapshot text = snapshot.Texts.First(x => x.Name == objectName);
        Paragraph paragraph = new();
        ParagraphProperties properties = new(new SpacingBetweenLines
        {
            Before = "0",
            After = "0"
        });

        if (TryGetJustification(text.HorzAlign, out JustificationValues justification))
            properties.Append(new Justification { Val = justification });

        paragraph.Append(properties);

        Run run = new();
        run.Append(CreateRunProperties(text));

        string[] lines = (text.Text ?? string.Empty)
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Split('\n');

        for (int i = 0; i < lines.Length; i++)
        {
            if (i > 0)
                run.Append(new Break());

            if (lines[i].Length > 0)
            {
                run.Append(new Text(lines[i])
                {
                    Space = SpaceProcessingModeValues.Preserve
                });
            }
        }

        paragraph.Append(run);
        return paragraph;
    }

    private static void ApplyCellStyle(TableCellProperties properties, DocxLayoutSnapshot snapshot, string objectName)
    {
        DocxTextSnapshot text = snapshot.Texts.First(x => x.Name == objectName);

        if (!string.IsNullOrEmpty(text.FillColor))
        {
            properties.Append(new Shading
            {
                Val = ShadingPatternValues.Clear,
                Fill = text.FillColor
            });
        }

        if (!string.IsNullOrEmpty(text.BorderColor))
        {
            properties.Append(new TableCellBorders(
                CreateBorder<TopBorder>(text.BorderColor),
                CreateBorder<LeftBorder>(text.BorderColor),
                CreateBorder<BottomBorder>(text.BorderColor),
                CreateBorder<RightBorder>(text.BorderColor)));
        }
    }

    private static T CreateBorder<T>(string color)
        where T : BorderType, new()
    {
        return new T
        {
            Val = BorderValues.Single,
            Color = color,
            Size = 8
        };
    }

    private static RunProperties CreateRunProperties(DocxTextSnapshot text)
    {
        RunProperties properties = new();

        if (!string.IsNullOrEmpty(text.FontName))
        {
            properties.Append(new RunFonts
            {
                Ascii = text.FontName,
                HighAnsi = text.FontName
            });
        }

        if (text.FontSizeInPoints > 0)
        {
            properties.Append(new FontSize
            {
                Val = ((int)Math.Round(text.FontSizeInPoints * 2, MidpointRounding.AwayFromZero)).ToString()
            });
        }

        if (text.Bold)
            properties.Append(new Bold());

        if (text.Italic)
            properties.Append(new Italic());

        if (text.Underline)
            properties.Append(new Underline { Val = UnderlineValues.Single });

        if (!string.IsNullOrEmpty(text.TextColor))
            properties.Append(new Color { Val = text.TextColor });

        return properties;
    }

    private static bool TryGetJustification(HorzAlign align, out JustificationValues value)
    {
        value = align switch
        {
            HorzAlign.Left => JustificationValues.Left,
            HorzAlign.Center => JustificationValues.Center,
            HorzAlign.Right => JustificationValues.Right,
            HorzAlign.Justify => JustificationValues.Both,
            _ => JustificationValues.Left
        };

        return align is HorzAlign.Left or HorzAlign.Center or HorzAlign.Right or HorzAlign.Justify;
    }

    private static int ToTwips(float pixels)
    {
        return Math.Max(1, DocxUnitConverter.PixelsToTwips(pixels));
    }
}
