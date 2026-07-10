namespace WordApi.Services.PdfConversion;

/// <summary>
/// A backend, in-process docx → PDF conversion engine. Implementations must be self-contained
/// (no external process / sidecar) so the .docx never leaves the machine — same privacy contract
/// as the existing <c>/api/document</c> endpoints.
/// </summary>
public interface IPdfConverter
{
    /// <summary>Engine id used on the wire (<c>?engine=…</c>), e.g. "oss" or "devexpress".</summary>
    string Engine { get; }

    /// <summary>Convert the .docx at <paramref name="docxPath"/> to a PDF byte array.</summary>
    byte[] Convert(string docxPath);
    byte[] Convert(Stream stream);
}
