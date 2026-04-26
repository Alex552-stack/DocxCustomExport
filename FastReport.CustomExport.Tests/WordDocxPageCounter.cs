using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace FastReport.CustomExport.Tests;

internal static class WordDocxPageCounter
{
    private const int WdStatisticPages = 2;
    private const int WdDoNotSaveChanges = 0;
    private const int WdAlertsNone = 0;

    public static int CountPages(Stream docxStream)
    {
        if (!OperatingSystem.IsWindows())
            throw new PlatformNotSupportedException("Word-backed DOCX page counting requires Windows.");

        string tempPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.docx");
        try
        {
            if (docxStream.CanSeek)
                docxStream.Position = 0;

            using (FileStream file = File.Create(tempPath))
                docxStream.CopyTo(file);

            return RunInStaThread(() => CountPagesCore(tempPath));
        }
        finally
        {
            if (File.Exists(tempPath))
                File.Delete(tempPath);
        }
    }

    private static int CountPagesCore(string path)
    {
        Type? wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType == null)
            throw new InvalidOperationException("Microsoft Word COM automation is not available.");

        object? wordApplication = null;
        object? document = null;

        try
        {
            wordApplication = Activator.CreateInstance(wordType)
                ?? throw new InvalidOperationException("Microsoft Word COM automation could not be started.");

            wordType.InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, wordApplication, new object[] { false });
            wordType.InvokeMember("DisplayAlerts", System.Reflection.BindingFlags.SetProperty, null, wordApplication, new object[] { WdAlertsNone });

            object documents = wordType.InvokeMember("Documents", System.Reflection.BindingFlags.GetProperty, null, wordApplication, null)!;
            object[] openArgs = new object[]
            {
                path,
                false,
                true
            };
            document = documents.GetType().InvokeMember("Open", System.Reflection.BindingFlags.InvokeMethod, null, documents, openArgs);
            document!.GetType().InvokeMember("Repaginate", System.Reflection.BindingFlags.InvokeMethod, null, document, null);

            object result = document.GetType().InvokeMember(
                "ComputeStatistics",
                System.Reflection.BindingFlags.InvokeMethod,
                null,
                document,
                new object[] { WdStatisticPages });

            return Convert.ToInt32(result);
        }
        finally
        {
            TryInvoke(document, "Close", WdDoNotSaveChanges);
            TryInvoke(wordApplication, "Quit", WdDoNotSaveChanges);
            ReleaseComObject(document);
            ReleaseComObject(wordApplication);
        }
    }

    private static int RunInStaThread(Func<int> action)
    {
        int result = 0;
        Exception? error = null;
        using ManualResetEvent finished = new(false);

        Thread thread = new(() =>
        {
            try
            {
                result = action();
            }
            catch (Exception ex)
            {
                error = ex;
            }
            finally
            {
                finished.Set();
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        finished.WaitOne();

        if (error != null)
            throw error;

        return result;
    }

    private static void TryInvoke(object? target, string methodName, params object[] args)
    {
        if (target == null)
            return;

        try
        {
            target.GetType().InvokeMember(methodName, System.Reflection.BindingFlags.InvokeMethod, null, target, args);
        }
        catch
        {
        }
    }

    private static void ReleaseComObject(object? value)
    {
        if (value != null && Marshal.IsComObject(value))
            Marshal.FinalReleaseComObject(value);
    }
}
