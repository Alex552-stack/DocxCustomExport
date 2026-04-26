using FastReport.Utils;

namespace FastReport.Export.Custom;

/// <summary>
/// Registers the custom export when the assembly is loaded as a FastReport plugin.
/// </summary>
public sealed class AssemblyInitializer : AssemblyInitializerBase
{
    /// <summary>
    /// Initializes a new instance of the <see cref="AssemblyInitializer"/> class.
    /// </summary>
    public AssemblyInitializer()
    {
#pragma warning disable CS0612
        RegisteredObjects.AddExport(typeof(DocxExport), "Word 2007", "DOCX export");
        RegisteredObjects.AddExport(typeof(EmptyCustomExport), "Custom", "Empty custom export");
#pragma warning restore CS0612
    }
}
