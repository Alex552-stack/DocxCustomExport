using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using WordTable = DocumentFormat.OpenXml.Wordprocessing.Table;
using A = DocumentFormat.OpenXml.Drawing;
using Drawing = DocumentFormat.OpenXml.Drawing;
using DrawingWordprocessing = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using DrawingFont = System.Drawing.Font;
using DrawingFontStyle = System.Drawing.FontStyle;
using DrawingGraphicsUnit = System.Drawing.GraphicsUnit;

namespace FastReport.Export.Custom;

internal static class DocxMatrixTableBuilder
{
    private static int nextDrawingId;

    public static WordTable Build(
        DocxMatrixPage page,
        OpenXmlPartContainer? owningPart = null,
        MainDocumentPart? mainDocumentPart = null)
    {
        mainDocumentPart ??= owningPart as MainDocumentPart;
        WordTable table = new();
        table.Append(BuildProperties(page));
        table.Append(BuildGrid(page));

        Dictionary<(int Row, int Column), DocxMatrixFragment> origins = page.Fragments
            .ToDictionary(x => (x.Row, x.Column));

        for (int rowIndex = 0; rowIndex < page.RowCount; rowIndex++)
        {
            TableRow row = new();
            row.Append(new TableRowProperties(
                new TableRowHeight
                {
                    Val = (UInt32)Math.Min(31680, ToTwips(page.YEdges[rowIndex + 1] - page.YEdges[rowIndex])),
                    HeightType = HeightRuleValues.Exact
                }));

            int columnIndex = 0;
            while (columnIndex < page.ColumnCount)
            {
                if (origins.TryGetValue((rowIndex, columnIndex), out DocxMatrixFragment? origin))
                {
                    row.Append(CreateOriginCell(page, origin, owningPart, mainDocumentPart));
                    columnIndex += origin.ColumnSpan;
                    continue;
                }

                DocxMatrixFragment? continuation = page.Fragments.FirstOrDefault(x =>
                    x.Column == columnIndex &&
                    x.Row < rowIndex &&
                    x.Row + x.RowSpan > rowIndex);

                if (continuation != null)
                {
                    row.Append(CreateContinuationCell(page, continuation));
                    columnIndex += continuation.ColumnSpan;
                    continue;
                }

                row.Append(CreateEmptyCell(page, columnIndex));
                columnIndex++;
            }

            table.Append(row);
        }

        return table;
    }

    private static TableProperties BuildProperties(DocxMatrixPage page)
    {
        return new TableProperties(
            new TableWidth
            {
                Type = TableWidthUnitValues.Dxa,
                Width = ToTwips(page.XEdges[^1] - page.XEdges[0]).ToString()
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

    private static TableGrid BuildGrid(DocxMatrixPage page)
    {
        TableGrid tableGrid = new();
        for (int index = 0; index < page.ColumnCount; index++)
        {
            tableGrid.Append(new GridColumn
            {
                Width = ToTwips(page.XEdges[index + 1] - page.XEdges[index]).ToString()
            });
        }

        return tableGrid;
    }

    private static TableCell CreateOriginCell(
        DocxMatrixPage page,
        DocxMatrixFragment origin,
        OpenXmlPartContainer? owningPart,
        MainDocumentPart? mainDocumentPart)
    {
        DocxMatrixObject source = page.Objects[origin.ObjectIndex];
        TableCell cell = new();
        TableCellProperties properties = CreateCellProperties(page, origin.Column, origin.ColumnSpan);

        if (origin.ColumnSpan > 1)
            properties.Append(new GridSpan { Val = origin.ColumnSpan });

        if (origin.RowSpan > 1)
            properties.Append(new VerticalMerge { Val = MergedCellValues.Restart });

        ApplyStyle(properties, source.Style);

        cell.Append(properties);

        if (source.IsText)
        {
            foreach (Paragraph paragraph in CreateParagraphs(source, mainDocumentPart))
                cell.Append(paragraph);
        }
        else if (source.HasImage)
        {
            cell.Append(CreateImageParagraph(source, owningPart));
        }
        else
        {
            cell.Append(new Paragraph());
        }

        return cell;
    }

    private static Paragraph CreateImageParagraph(DocxMatrixObject source, OpenXmlPartContainer? owningPart)
    {
        Paragraph paragraph = new(new ParagraphProperties(
            new SpacingBetweenLines
            {
                Before = "0",
                After = "0"
            }));

        if (owningPart == null || !source.HasImage)
        {
            paragraph.Append(new Run());
            return paragraph;
        }

        ImagePart? imagePart = AddImagePart(owningPart);
        if (imagePart == null)
        {
            paragraph.Append(new Run());
            return paragraph;
        }

        using (System.IO.MemoryStream stream = new(source.ImageBytes!))
            imagePart.FeedData(stream);

        string relationshipId = owningPart.GetIdOfPart(imagePart);
        Run run = new(new DocumentFormat.OpenXml.Wordprocessing.Drawing(
            new DrawingWordprocessing.Inline(
                new DrawingWordprocessing.Extent
                {
                    Cx = DocxUnitConverter.PixelsToEmus(Math.Max(1, source.ImageWidth)),
                    Cy = DocxUnitConverter.PixelsToEmus(Math.Max(1, source.ImageHeight))
                },
                new DrawingWordprocessing.EffectExtent
                {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L
                },
                new DrawingWordprocessing.DocProperties
                {
                    Id = (UInt32)Interlocked.Increment(ref nextDrawingId),
                    Name = string.IsNullOrWhiteSpace(source.Name) ? "Picture" : source.Name
                },
                new DrawingWordprocessing.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks
                    {
                        NoChangeAspect = true
                    }),
                new A.Graphic(
                    new A.GraphicData(
                        new Pic.Picture(
                            new Pic.NonVisualPictureProperties(
                                new Pic.NonVisualDrawingProperties
                                {
                                    Id = 0U,
                                    Name = string.IsNullOrWhiteSpace(source.Name) ? "Picture.png" : $"{source.Name}.png"
                                },
                                new Pic.NonVisualPictureDrawingProperties()),
                            new Pic.BlipFill(
                                new A.Blip
                                {
                                    Embed = relationshipId
                                },
                                new A.Stretch(new A.FillRectangle())),
                            new Pic.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0L, Y = 0L },
                                    new A.Extents
                                    {
                                        Cx = DocxUnitConverter.PixelsToEmus(Math.Max(1, source.ImageWidth)),
                                        Cy = DocxUnitConverter.PixelsToEmus(Math.Max(1, source.ImageHeight))
                                    }),
                                new A.PresetGeometry(new A.AdjustValueList())
                                {
                                    Preset = A.ShapeTypeValues.Rectangle
                                })))
                    {
                        Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                    }))
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            }));

        paragraph.Append(run);
        return paragraph;
    }

    private static ImagePart? AddImagePart(OpenXmlPartContainer owningPart)
    {
        return owningPart switch
        {
            MainDocumentPart mainDocumentPart => mainDocumentPart.AddImagePart(ImagePartType.Png),
            HeaderPart headerPart => headerPart.AddImagePart(ImagePartType.Png),
            FooterPart footerPart => footerPart.AddImagePart(ImagePartType.Png),
            _ => null
        };
    }

    private static TableCell CreateContinuationCell(DocxMatrixPage page, DocxMatrixFragment continuation)
    {
        DocxMatrixObject source = page.Objects[continuation.ObjectIndex];
        TableCell cell = new();
        TableCellProperties properties = CreateCellProperties(page, continuation.Column, continuation.ColumnSpan);
        if (continuation.ColumnSpan > 1)
            properties.Append(new GridSpan { Val = continuation.ColumnSpan });

        properties.Append(new VerticalMerge());
        ApplyStyle(properties, source.Style);
        cell.Append(properties);
        cell.Append(new Paragraph());
        return cell;
    }

    private static TableCell CreateEmptyCell(DocxMatrixPage page, int columnIndex)
    {
        TableCell cell = new();
        cell.Append(CreateCellProperties(page, columnIndex, 1));
        cell.Append(new Paragraph());
        return cell;
    }

    private static TableCellProperties CreateCellProperties(DocxMatrixPage page, int columnIndex, int columnSpan)
    {
        int width = ToTwips(page.XEdges[columnIndex + columnSpan] - page.XEdges[columnIndex]);
        return new TableCellProperties(
            new TableCellWidth
            {
                Type = TableWidthUnitValues.Dxa,
                Width = width.ToString()
            });
    }

    private static void ApplyStyle(TableCellProperties properties, DocxMatrixStyle style)
    {
        if (style.HasBorder)
        {
            properties.Append(new TableCellBorders(
                CreateBorder<TopBorder>(style.BorderColor),
                CreateBorder<LeftBorder>(style.BorderColor),
                CreateBorder<BottomBorder>(style.BorderColor),
                CreateBorder<RightBorder>(style.BorderColor)));
        }

        if (style.HasFill)
        {
            properties.Append(new Shading
            {
                Val = ShadingPatternValues.Clear,
                Fill = style.FillColor
            });
        }

        if (style.PaddingLeft > 0 || style.PaddingTop > 0 || style.PaddingRight > 0 || style.PaddingBottom > 0)
        {
            properties.Append(new TableCellMargin(
                new TopMargin
                {
                    Width = ToTwips(style.PaddingTop).ToString(),
                    Type = TableWidthUnitValues.Dxa
                },
                new LeftMargin
                {
                    Width = ToTwips(style.PaddingLeft).ToString(),
                    Type = TableWidthUnitValues.Dxa
                },
                new BottomMargin
                {
                    Width = ToTwips(style.PaddingBottom).ToString(),
                    Type = TableWidthUnitValues.Dxa
                },
                new RightMargin
                {
                    Width = ToTwips(style.PaddingRight).ToString(),
                    Type = TableWidthUnitValues.Dxa
                }));
        }

        if (style.VertAlign != VertAlign.Top)
        {
            properties.Append(new TableCellVerticalAlignment
            {
                Val = style.VertAlign switch
                {
                    VertAlign.Center => TableVerticalAlignmentValues.Center,
                    VertAlign.Bottom => TableVerticalAlignmentValues.Bottom,
                    _ => TableVerticalAlignmentValues.Top
                }
            });
        }
    }

    private static IEnumerable<Paragraph> CreateParagraphs(DocxMatrixObject source, MainDocumentPart? mainDocumentPart)
    {
        string[] logicalParagraphs = (source.Text ?? string.Empty)
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Split('\n');

        if (logicalParagraphs.Length == 0)
        {
            yield return new Paragraph(new Run());
            yield break;
        }

        Paragraph paragraph = new();
        ParagraphProperties properties = new(CreateSpacing(source.Style));

        if (source.Style.ParagraphOffset > 0)
        {
            properties.Append(new Indentation
            {
                FirstLine = ToTwips(source.Style.ParagraphOffset).ToString()
            });
        }

        if (TryGetJustification(source.Style.HorzAlign, out JustificationValues justification))
            properties.Append(new Justification { Val = justification });

        if (source.Style.RightToLeft)
            properties.Append(new BiDi());

        paragraph.Append(properties);

        if (!string.IsNullOrWhiteSpace(source.Bookmark))
        {
            paragraph.Append(new BookmarkStart
            {
                Name = source.Bookmark,
                Id = "0"
            });
        }

        OpenXmlElement content = CreateRunContainer(source, logicalParagraphs, mainDocumentPart);
        paragraph.Append(content);

        if (!string.IsNullOrWhiteSpace(source.Bookmark))
            paragraph.Append(new BookmarkEnd { Id = "0" });

        yield return paragraph;
    }

    private static OpenXmlElement CreateRunContainer(
        DocxMatrixObject source,
        string[] logicalParagraphs,
        MainDocumentPart? mainDocumentPart)
    {
        Run run = new();
        run.Append(CreateRunProperties(source.Style));

        for (int index = 0; index < logicalParagraphs.Length; index++)
        {
            if (index > 0)
                run.Append(new Break());

            if (logicalParagraphs[index].Length > 0)
            {
                run.Append(new Text(logicalParagraphs[index])
                {
                    Space = SpaceProcessingModeValues.Preserve
                });
            }
        }

        if (source.HyperlinkKind == HyperlinkKind.URL && !string.IsNullOrWhiteSpace(source.Hyperlink) && mainDocumentPart != null)
        {
            HyperlinkRelationship relationship = mainDocumentPart.AddHyperlinkRelationship(
                new Uri(source.Hyperlink, UriKind.Absolute),
                true);

            DocumentFormat.OpenXml.Wordprocessing.Hyperlink hyperlink = new()
            {
                Id = relationship.Id,
                History = OnOffValue.FromBoolean(true)
            };
            hyperlink.Append(run);
            return hyperlink;
        }

        return run;
    }

    private static RunProperties CreateRunProperties(DocxMatrixStyle style)
    {
        RunProperties properties = new();

        if (!string.IsNullOrEmpty(style.FontName))
        {
            properties.Append(new RunFonts
            {
                Ascii = style.FontName,
                HighAnsi = style.FontName
            });
        }

        if (style.Bold)
            properties.Append(new Bold());

        if (style.Italic)
            properties.Append(new Italic());

        if (!string.IsNullOrEmpty(style.TextColor))
            properties.Append(new DocumentFormat.OpenXml.Wordprocessing.Color { Val = style.TextColor });

        if (style.FontSizeInPoints > 0)
        {
            string halfPoints = ((int)Math.Round(style.FontSizeInPoints * 2, MidpointRounding.AwayFromZero)).ToString();
            properties.Append(
                new FontSize { Val = halfPoints },
                new FontSizeComplexScript { Val = halfPoints });
        }

        if (style.Underline)
            properties.Append(new Underline { Val = UnderlineValues.Single });

        return properties;
    }

    private static SpacingBetweenLines CreateSpacing(DocxMatrixStyle style)
    {
        int line = GetLineHeightTwips(style);
        if (line <= 0)
        {
            return new SpacingBetweenLines
            {
                Before = "0",
                After = "0"
            };
        }

        return new SpacingBetweenLines
        {
            Before = "0",
            After = "0",
            Line = line.ToString(),
            LineRule = LineSpacingRuleValues.Exact
        };
    }

    private static int GetLineHeightTwips(DocxMatrixStyle style)
    {
        if (style.LineHeight > 0)
            return Math.Max(1, DocxUnitConverter.PixelsToTwips(style.LineHeight));

        if (string.IsNullOrWhiteSpace(style.FontName) || style.FontSizeInPoints <= 0)
            return 0;

        DrawingFontStyle fontStyle = DrawingFontStyle.Regular;
        if (style.Bold)
            fontStyle |= DrawingFontStyle.Bold;
        if (style.Italic)
            fontStyle |= DrawingFontStyle.Italic;

        using DrawingFont font = new(style.FontName, style.FontSizeInPoints, fontStyle, DrawingGraphicsUnit.Point);
        float pixels = font.GetHeight();
        return Math.Max(1, DocxUnitConverter.PixelsToTwips(pixels));
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

    private static int ToTwips(float pixels)
    {
        return Math.Max(1, DocxUnitConverter.PixelsToTwips(pixels));
    }
}
