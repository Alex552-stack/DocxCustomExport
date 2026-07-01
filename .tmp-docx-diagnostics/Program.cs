using System.Reflection;
using FastReport;

using Report report = new();
report.LoadPrepared(Path.GetFullPath("Tools/FastReport.CustomExport.Tests/TestReports/AdevGenerala.fpx"));
Type builderType = typeof(FastReport.Export.Custom.DocxExport).Assembly.GetType("FastReport.Export.Custom.DocxMatrixBuilder")!;
MethodInfo buildRegions = builderType.GetMethod("BuildRegions", BindingFlags.Public | BindingFlags.Static)!;
foreach (int i in new[] {4,12,13,18,22})
{
    using ReportPage page = report.PreparedPages.GetPage(i);
    object regions = buildRegions.Invoke(null, new object[] { page, 0.5f })!;
    object body = regions.GetType().GetProperty("Body")!.GetValue(regions)!;
    var objects = (System.Collections.IEnumerable)body.GetType().GetProperty("Objects")!.GetValue(body)!;
    Console.WriteLine($"PAGE {i}: landscape={page.Landscape} sizePx={page.WidthInPixels:0.##}");
    var rows = new List<(string Name, string Kind, float Left, float Width, float Right, float Top, string Text)>();
    foreach (object obj in objects)
    {
        Type t = obj.GetType();
        string name = (string)t.GetProperty("Name")!.GetValue(obj)!;
        string kind = (string)t.GetProperty("Kind")!.GetValue(obj)!;
        float left = (float)t.GetProperty("Left")!.GetValue(obj)!;
        float width = (float)t.GetProperty("Width")!.GetValue(obj)!;
        float top = (float)t.GetProperty("Top")!.GetValue(obj)!;
        string text = ((string)t.GetProperty("Text")!.GetValue(obj)!).Replace("\r", " ").Replace("\n", " ");
        if (left + width > page.WidthInPixels + 50)
            rows.Add((name, kind, left, width, left + width, top, text));
    }
    foreach (var r in rows.OrderByDescending(x => x.Right).Take(12))
    {
        string text = r.Text.Substring(0, Math.Min(70, r.Text.Length));
        Console.WriteLine($"  {r.Kind} {r.Name} left={r.Left:0.##} width={r.Width:0.##} right={r.Right:0.##} top={r.Top:0.##} text='{text}'");
    }
}
