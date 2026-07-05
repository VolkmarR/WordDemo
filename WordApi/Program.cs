using WordApi.Models;
using WordApi.Services;
using WordApi.Services.PdfConversion;

var builder = WebApplication.CreateBuilder(args);

// QuestPDF is dual-licensed; a paid license is required above $1M annual revenue. Acknowledge it
// once at startup. Use LicenseType.Enterprise if more than 10 developers reference QuestPDF.
QuestPDF.Settings.License = QuestPDF.Infrastructure.LicenseType.Professional;

// In-process docx → PDF engines, keyed by their wire id (?engine=…).
var pdfConverters = new IPdfConverter[]
    {
        new QuestPdfConverter(), new MigraDocConverter(), new DevExpressPdfConverter(),
    }
    .ToDictionary(c => c.Engine, StringComparer.OrdinalIgnoreCase);

// OpenAPI document (served at /openapi/v1.json) + Swagger UI on top of it.
builder.Services.AddOpenApi();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.MapOpenApi();
    app.UseSwaggerUI(o =>
    {
        o.SwaggerEndpoint("/openapi/v1.json", "WordApi v1");
        o.DocumentTitle = "WordApi — Swagger UI";
    });
}

app.UseHttpsRedirection();

// Resolve the .docx path from config key "WordFile" (default ../word-demo.docx),
// relative to the content root. The file never leaves the machine.
string ResolveWordPath() =>
    Path.GetFullPath(Path.Combine(
        app.Environment.ContentRootPath,
        app.Configuration["WordFile"] ?? "../word-demo.docx"));

const string DocxContentType =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

// Approach A: parsed on the server (Open XML SDK) → shared DocumentModel JSON.
app.MapGet("/api/document", () =>
    {
        var path = ResolveWordPath();
        if (!File.Exists(path)) return Results.NotFound($"Word file not found: {path}");
        return Results.Ok(DocumentReader.Read(path));
    })
    .WithName("GetDocument")
    .WithTags("Document")
    .WithSummary("Parsed document model")
    .WithDescription("Reads the .docx on the server with the Open XML SDK and returns a flat, "
                     + "style-resolved DocumentModel as JSON (approach A).")
    .Produces<DocumentModel>()
    .Produces(StatusCodes.Status404NotFound);

// Approaches B, C, D: raw .docx bytes for browser-side parsing. Nothing leaves the browser.
app.MapGet("/api/document/file", () =>
    {
        var path = ResolveWordPath();
        if (!File.Exists(path)) return Results.NotFound($"Word file not found: {path}");
        var bytes = File.ReadAllBytes(path);
        return Results.File(bytes, DocxContentType, Path.GetFileName(path));
    })
    .WithName("GetDocumentFile")
    .WithTags("Document")
    .WithSummary("Raw .docx bytes")
    .WithDescription("Serves the original .docx same-origin for the browser-side renderers "
                     + "(docx-preview, mammoth).")
    .Produces(StatusCodes.Status200OK, contentType: DocxContentType)
    .Produces(StatusCodes.Status404NotFound);

// Approach E / F: convert the .docx to PDF on the server, in-process, and stream it back inline.
// ?engine=oss (QuestPDF, default) or ?engine=devexpress. The .docx never leaves the machine.
app.MapGet("/api/document/pdf", (string? engine) =>
    {
        var key = string.IsNullOrWhiteSpace(engine) ? "oss" : engine;
        if (!pdfConverters.TryGetValue(key, out var converter))
            return Results.BadRequest(
                $"Unknown engine '{engine}'. Use one of: {string.Join(", ", pdfConverters.Keys)}.");

        var path = ResolveWordPath();
        if (!File.Exists(path)) return Results.NotFound($"Word file not found: {path}");

        var started = System.Diagnostics.Stopwatch.GetTimestamp();
        byte[] pdf;
        try
        {
            pdf = converter.Convert(path);
        }
        catch (NotSupportedException ex)
        {
            // e.g. the DevExpress engine isn't configured in this build.
            return Results.Problem(title: "PDF engine unavailable", detail: ex.Message,
                statusCode: StatusCodes.Status501NotImplemented);
        }
        var ms = System.Diagnostics.Stopwatch.GetElapsedTime(started).TotalMilliseconds;

        var response = Results.File(pdf, "application/pdf", $"word-demo-{key}.pdf");
        return new HeaderResult(response, ("X-Convert-Ms", ((int)ms).ToString()),
                                          ("X-Pdf-Bytes", pdf.Length.ToString()),
                                          ("Access-Control-Expose-Headers", "X-Convert-Ms, X-Pdf-Bytes"));
    })
    .WithName("GetDocumentPdf")
    .WithTags("Document")
    .WithSummary("Server-rendered PDF")
    .WithDescription("Converts the .docx to PDF in-process and returns it inline. "
                     + "engine=oss (QuestPDF, default) | migradoc (PDFsharp/MigraDoc, free) "
                     + "| devexpress (DevExpress Office File API).")
    .Produces(StatusCodes.Status200OK, contentType: "application/pdf")
    .Produces(StatusCodes.Status400BadRequest)
    .Produces(StatusCodes.Status404NotFound)
    .Produces(StatusCodes.Status501NotImplemented);

app.Run();

// Wraps an IResult to add response headers (conversion time + PDF size for the UI status line).
sealed record HeaderResult(IResult Inner, params (string Name, string Value)[] Headers) : IResult
{
    public Task ExecuteAsync(HttpContext ctx)
    {
        foreach (var (name, value) in Headers) ctx.Response.Headers[name] = value;
        return Inner.ExecuteAsync(ctx);
    }
}
