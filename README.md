# FastReport.CustomExport

Custom DOCX export implementation for FastReport targeting .NET 8.

## Projects

- `FastReport.CustomExport` - export library.
- `FastReport.CustomExport.Tests` - focused tests and local report fixtures.

The library depends on the `FastReport.OpenSource` NuGet package, not on a sibling FastReport source checkout.

## Build and Test

Requires .NET SDK 10.0.300 or newer for `.slnx` tooling. The package itself targets `net8.0`.

```powershell
dotnet test FastReport.CustomExport.slnx
```

## Local NuGet Feed

Create or update the local package:

```powershell
dotnet pack FastReport.CustomExport/FastReport.CustomExport.csproj -c Release
```

Use the generated folder as a NuGet source in a consuming project:

```powershell
dotnet nuget add source D:\Tools\DocxExport\DocxCustomExport\.artifacts\packages -n LocalDocxCustomExport
dotnet add path\to\YourApp.csproj package FastReport.CustomExport --version 0.1.1-local
```

If the package already exists in the consuming app, clear NuGet caches before reinstalling the same local version:

```powershell
dotnet nuget locals all --clear
```
