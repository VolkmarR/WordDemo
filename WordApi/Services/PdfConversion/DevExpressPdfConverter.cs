namespace WordApi.Services.PdfConversion;

/// <summary>
/// Approach F (full commercial). Converts the raw .docx to PDF with the DevExpress Office &amp; PDF
/// File API in a single call — <c>RichEditDocumentServer.ExportToPdf</c>. Runs fully in-process and
/// is Linux-clean via the <c>DevExpress.Drawing.Skia</c> engine (SkiaSharp; no libgdiplus).
///
/// PREREQUISITE (not wired by default): DevExpress packages are NOT on nuget.org. To enable:
///   1. Add the DevExpress NuGet feed (with your account auth key) to a nuget.config.
///   2. Reference DevExpress.Document.Processor + DevExpress.Drawing.Skia in WordApi.csproj.
///   3. Register your DevExpress license, and define the DEVEXPRESS build symbol
///      (&lt;DefineConstants&gt;$(DefineConstants);DEVEXPRESS&lt;/DefineConstants&gt; in the .csproj).
///
/// Until then this compiles as a stub that reports the engine is not configured, so the project
/// builds out of the box. If a DevExpress license is unavailable, GemBox.Document (on nuget.org,
/// emailed trial key) is a drop-in alternative implementing this same interface.
/// </summary>
public sealed class DevExpressPdfConverter : IPdfConverter
{
    public string Engine => "devexpress";

    public byte[] Convert(string docxPath)
    {
#if DEVEXPRESS
        using var server = new DevExpress.XtraRichEdit.RichEditDocumentServer();
        server.LoadDocument(docxPath, DevExpress.XtraRichEdit.DocumentFormat.OpenXml);
        using var ms = new MemoryStream();
        server.ExportToPdf(ms);
        return ms.ToArray();
#else
        throw new NotSupportedException(
            "The DevExpress PDF engine is not configured in this build. Add the DevExpress NuGet "
            + "feed + license and define the DEVEXPRESS build symbol (see DevExpressPdfConverter "
            + "for steps), or use ?engine=oss (QuestPDF).");
#endif
    }
}
