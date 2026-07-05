using WordApi.Models;
using WordApi.Services;

var builder = WebApplication.CreateBuilder(args);

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

app.Run();
