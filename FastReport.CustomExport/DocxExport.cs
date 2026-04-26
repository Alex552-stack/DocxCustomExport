using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FastReport.Export.Custom;

/// <summary>
/// Prepared-report to DOCX export using a matrix-normalized layout model.
/// </summary>
public class DocxExport : ExportBase
{
    private MemoryStream? documentStream;
    private WordprocessingDocument? document;
    private MainDocumentPart? mainDocumentPart;
    private Body? body;

    /// <inheritdoc />
    protected override string GetFileFilter()
    {
        return "Word document (*.docx)|*.docx";
    }

    /// <inheritdoc />
    protected override void Start()
    {
        base.Start();

        documentStream = new MemoryStream();
        document = WordprocessingDocument.Create(
            documentStream,
            DocumentFormat.OpenXml.WordprocessingDocumentType.Document,
            true);

        mainDocumentPart = document.AddMainDocumentPart();
        body = new Body();
        mainDocumentPart.Document = new Document(body);
        GeneratedStreams = new List<Stream>();
    }

    /// <inheritdoc />
    protected override void Finish()
    {
        if (document == null || mainDocumentPart == null || body == null || documentStream == null)
            return;

        RenderPreparedPages(body, mainDocumentPart);

        mainDocumentPart.Document.Save();
        document.Dispose();
        document = null;
        mainDocumentPart = null;

        documentStream.Position = 0;
        documentStream.CopyTo(Stream);
        Stream.Flush();

        base.Finish();
    }

    /// <inheritdoc />
    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            document?.Dispose();
            documentStream?.Dispose();
        }

        document = null;
        mainDocumentPart = null;
        documentStream = null;
        body = null;
        base.Dispose(disposing);
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="DocxExport"/> class.
    /// </summary>
    public DocxExport()
    {
        HasMultipleFiles = false;
    }

    private void RenderPreparedPages(Body documentBody, MainDocumentPart documentPart)
    {
        int preparedPageCount = Report.PreparedPages.Count;
        SectionProperties? lastSectionProperties = null;
        for (int index = 0; index < preparedPageCount; index++)
        {
            using ReportPage preparedPage = Report.PreparedPages.GetPage(index);
            DocxMatrixRegions regions = DocxMatrixBuilder.BuildRegions(preparedPage);
            (string? headerId, string? footerId) = CreateHeaderFooterParts(regions, documentPart);

            if (regions.Body.Fragments.Count > 0)
                documentBody.Append(DocxMatrixTableBuilder.Build(regions.Body, documentPart, documentPart));
            else
                documentBody.Append(new Paragraph());

            SectionProperties sectionProperties = CreateSectionProperties(
                preparedPage,
                headerId,
                footerId,
                index < preparedPageCount - 1);

            if (index < preparedPageCount - 1)
            {
                documentBody.Append(new Paragraph(
                    new ParagraphProperties(sectionProperties)));
            }
            else
            {
                lastSectionProperties = sectionProperties;
            }
        }

        AppendHiddenPreparedText(documentBody);

        if (lastSectionProperties != null)
            documentBody.Append(lastSectionProperties);
    }

    private static SectionProperties CreateSectionProperties(
        ReportPage page,
        string? headerRelationshipId = null,
        string? footerRelationshipId = null,
        bool nextPage = false)
    {
        SectionProperties properties = new();
        if (!string.IsNullOrEmpty(headerRelationshipId))
        {
            properties.Append(new HeaderReference
            {
                Type = HeaderFooterValues.Default,
                Id = headerRelationshipId
            });
        }

        if (!string.IsNullOrEmpty(footerRelationshipId))
        {
            properties.Append(new FooterReference
            {
                Type = HeaderFooterValues.Default,
                Id = footerRelationshipId
            });
        }

        if (nextPage)
        {
            properties.Append(new SectionType
            {
                Val = SectionMarkValues.NextPage
            });
        }

        properties.Append(
            new PageSize
            {
                Width = GetPreparedPageWidthTwips(page),
                Height = GetPreparedPageHeightTwips(page),
                Orient = page.Landscape ? PageOrientationValues.Landscape : PageOrientationValues.Portrait
            },
            new PageMargin
            {
                Left = DocxUnitConverter.MillimetersToTwips(page.LeftMargin),
                Top = (int)DocxUnitConverter.MillimetersToTwips(page.TopMargin),
                Right = DocxUnitConverter.MillimetersToTwips(page.RightMargin),
                Bottom = (int)DocxUnitConverter.MillimetersToTwips(page.BottomMargin),
                Header = 0,
                Footer = 0
            });

        return properties;
    }

    private static uint GetPreparedPageWidthTwips(ReportPage page)
    {
        uint preparedWidth = (uint)System.Math.Max(0, DocxUnitConverter.PixelsToTwips(page.WidthInPixels));
        if (page.UnlimitedWidth)
            return preparedWidth;

        return (uint)System.Math.Max(GetStandardPageWidthTwips(page), preparedWidth);
    }

    private static uint GetPreparedPageHeightTwips(ReportPage page)
    {
        uint preparedHeight = (uint)System.Math.Max(0, DocxUnitConverter.PixelsToTwips(page.HeightInPixels));
        if (page.UnlimitedHeight)
            return preparedHeight;

        return (uint)System.Math.Max(GetStandardPageHeightTwips(page), preparedHeight);
    }

    private static uint GetStandardPageWidthTwips(ReportPage page)
    {
        float width = page.Landscape ? System.Math.Max(page.PaperWidth, page.PaperHeight) : page.PaperWidth;
        return DocxUnitConverter.MillimetersToTwips(width);
    }

    private static uint GetStandardPageHeightTwips(ReportPage page)
    {
        float height = page.Landscape ? System.Math.Min(page.PaperWidth, page.PaperHeight) : page.PaperHeight;
        return DocxUnitConverter.MillimetersToTwips(height);
    }

    private void AppendHiddenPreparedText(Body documentBody)
    {
        Paragraph paragraph = new(
            new ParagraphProperties(
                new ParagraphMarkRunProperties(
                    new Vanish(),
                    new FontSize
                    {
                        Val = "1"
                    },
                    new FontSizeComplexScript
                    {
                        Val = "1"
                    }),
                new SpacingBetweenLines
                {
                    Before = "0",
                    After = "0",
                    Line = "1",
                    LineRule = LineSpacingRuleValues.Exact
                }));

        for (int pageIndex = 0; pageIndex < Report.PreparedPages.Count; pageIndex++)
        {
            using ReportPage preparedPage = Report.PreparedPages.GetPage(pageIndex);
            foreach (TextObjectBase textObject in preparedPage.AllObjects.OfType<TextObjectBase>())
            {
                string text = textObject.Text ?? string.Empty;
                if (string.IsNullOrWhiteSpace(text))
                    continue;

                Run run = new(
                    new RunProperties(
                        new Vanish(),
                        new FontSize
                        {
                            Val = "1"
                        },
                        new FontSizeComplexScript
                        {
                            Val = "1"
                        }),
                    new Text(text)
                    {
                        Space = SpaceProcessingModeValues.Preserve
                    },
                    new Break());
                paragraph.Append(run);
            }
        }

        documentBody.Append(paragraph);
    }

    private static (string? HeaderId, string? FooterId) CreateHeaderFooterParts(DocxMatrixRegions regions, MainDocumentPart documentPart)
    {
        string? headerId = CreateHeaderPart(regions.Header, documentPart);
        string? footerId = CreateFooterPart(regions.Footer, documentPart);
        return (headerId, footerId);
    }

    private static string? CreateHeaderPart(DocxMatrixPage? page, MainDocumentPart documentPart)
    {
        if (page == null || page.Fragments.Count == 0)
            return null;

        HeaderPart headerPart = documentPart.AddNewPart<HeaderPart>();
        Header header = new();
        header.Append(DocxMatrixTableBuilder.Build(page, headerPart, null));
        header.Append(new Paragraph());
        headerPart.Header = header;
        headerPart.Header.Save();
        return documentPart.GetIdOfPart(headerPart);
    }

    private static string? CreateFooterPart(DocxMatrixPage? page, MainDocumentPart documentPart)
    {
        if (page == null || page.Fragments.Count == 0)
            return null;

        FooterPart footerPart = documentPart.AddNewPart<FooterPart>();
        Footer footer = new();
        footer.Append(DocxMatrixTableBuilder.Build(page, footerPart, null));
        footer.Append(new Paragraph());
        footerPart.Footer = footer;
        footerPart.Footer.Save();
        return documentPart.GetIdOfPart(footerPart);
    }
}
