using System;
using System.Collections.Generic;
using System.Linq;

namespace FastReport.Export.Custom;

public sealed class DocxLayoutGrid
{
    public IReadOnlyList<float> XEdges { get; init; } = Array.Empty<float>();

    public IReadOnlyList<float> YEdges { get; init; } = Array.Empty<float>();

    public IReadOnlyList<DocxLayoutGridObject> Objects { get; init; } = Array.Empty<DocxLayoutGridObject>();

    public int ColumnCount => Math.Max(0, XEdges.Count - 1);

    public int RowCount => Math.Max(0, YEdges.Count - 1);
}

public sealed class DocxLayoutGridObject
{
    public string Name { get; init; } = string.Empty;

    public string Kind { get; init; } = string.Empty;

    public string ParentBandName { get; init; } = string.Empty;

    public float Left { get; init; }

    public float Top { get; init; }

    public float Width { get; init; }

    public float Height { get; init; }

    public int Column { get; init; }

    public int Row { get; init; }

    public int ColumnSpan { get; init; }

    public int RowSpan { get; init; }

    public bool HasFill { get; init; }

    public bool HasBorder { get; init; }
}

public static class DocxLayoutGridBuilder
{
    private const float DefaultTolerance = 0.5f;

    public static DocxLayoutGrid Build(DocxLayoutSnapshot snapshot, float tolerance = DefaultTolerance)
    {
        List<float> xEdges = BuildEdges(
            snapshot.PageWidth,
            tolerance,
            snapshot.Bands.SelectMany(ToHorizontalEdges)
                .Concat(snapshot.Texts.SelectMany(ToHorizontalEdges)));

        List<float> yEdges = BuildEdges(
            snapshot.PageHeight,
            tolerance,
            snapshot.Bands.SelectMany(ToVerticalEdges)
                .Concat(snapshot.Texts.SelectMany(ToVerticalEdges)));

        List<DocxLayoutGridObject> objects = new();
        objects.AddRange(snapshot.Bands.Select(band => CreateBandObject(band, xEdges, yEdges, tolerance)));
        objects.AddRange(snapshot.Texts.Select(text => CreateTextObject(text, xEdges, yEdges, tolerance)));

        return new DocxLayoutGrid
        {
            XEdges = xEdges,
            YEdges = yEdges,
            Objects = objects
                .OrderBy(x => x.Row)
                .ThenBy(x => x.Column)
                .ThenBy(x => x.Kind)
                .ToList()
        };
    }

    private static DocxLayoutGridObject CreateBandObject(
        DocxBandSnapshot band,
        IReadOnlyList<float> xEdges,
        IReadOnlyList<float> yEdges,
        float tolerance)
    {
        return new DocxLayoutGridObject
        {
            Name = band.Name,
            Kind = band.BandType,
            Left = band.Left,
            Top = band.Top,
            Width = band.Width,
            Height = band.Height,
            Column = FindEdgeIndex(xEdges, band.Left, tolerance),
            Row = FindEdgeIndex(yEdges, band.Top, tolerance),
            ColumnSpan = FindSpan(xEdges, band.Left, band.Width, tolerance),
            RowSpan = FindSpan(yEdges, band.Top, band.Height, tolerance),
            HasFill = band.HasFill,
            HasBorder = band.HasBorder
        };
    }

    private static DocxLayoutGridObject CreateTextObject(
        DocxTextSnapshot text,
        IReadOnlyList<float> xEdges,
        IReadOnlyList<float> yEdges,
        float tolerance)
    {
        return new DocxLayoutGridObject
        {
            Name = text.Name,
            Kind = nameof(TextObjectBase),
            ParentBandName = text.ParentBandName,
            Left = text.Left,
            Top = text.Top,
            Width = text.Width,
            Height = text.Height,
            Column = FindEdgeIndex(xEdges, text.Left, tolerance),
            Row = FindEdgeIndex(yEdges, text.Top, tolerance),
            ColumnSpan = FindSpan(xEdges, text.Left, text.Width, tolerance),
            RowSpan = FindSpan(yEdges, text.Top, text.Height, tolerance),
            HasFill = text.HasFill,
            HasBorder = text.HasBorder
        };
    }

    private static IEnumerable<float> ToHorizontalEdges(DocxBandSnapshot band)
    {
        yield return band.Left;
        yield return band.Left + band.Width;
    }

    private static IEnumerable<float> ToHorizontalEdges(DocxTextSnapshot text)
    {
        yield return text.Left;
        yield return text.Left + text.Width;
    }

    private static IEnumerable<float> ToVerticalEdges(DocxBandSnapshot band)
    {
        yield return band.Top;
        yield return band.Top + band.Height;
    }

    private static IEnumerable<float> ToVerticalEdges(DocxTextSnapshot text)
    {
        yield return text.Top;
        yield return text.Top + text.Height;
    }

    private static List<float> BuildEdges(float pageExtent, float tolerance, IEnumerable<float> edges)
    {
        List<float> ordered = edges
            .Append(0)
            .Append(pageExtent)
            .Where(x => x >= 0 && x <= pageExtent)
            .OrderBy(x => x)
            .ToList();

        List<float> merged = new();
        foreach (float edge in ordered)
        {
            if (merged.Count == 0 || Math.Abs(edge - merged[^1]) > tolerance)
                merged.Add(edge);
        }

        if (merged.Count == 0 || merged[0] > tolerance)
            merged.Insert(0, 0);

        if (Math.Abs(merged[^1] - pageExtent) > tolerance)
            merged.Add(pageExtent);

        return merged;
    }

    private static int FindEdgeIndex(IReadOnlyList<float> edges, float value, float tolerance)
    {
        for (int i = 0; i < edges.Count; i++)
        {
            if (Math.Abs(edges[i] - value) <= tolerance)
                return i;
        }

        throw new InvalidOperationException($"Edge {value} was not found in the layout grid.");
    }

    private static int FindSpan(IReadOnlyList<float> edges, float start, float extent, float tolerance)
    {
        float end = start + extent;
        int startIndex = FindEdgeIndex(edges, start, tolerance);
        int endIndex = FindEdgeIndex(edges, end, tolerance);
        return Math.Max(1, endIndex - startIndex);
    }
}
