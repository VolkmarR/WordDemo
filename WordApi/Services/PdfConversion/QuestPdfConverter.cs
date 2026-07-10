using QuestPDF.Fluent;
using QuestPDF.Infrastructure;
using WordApi.Models;

namespace WordApi.Services.PdfConversion;

/// <summary>
/// Approach E (low-cost). Renders the same style-resolved <see cref="DocumentModel"/> that
/// approach A's custom React renderer consumes (produced by <see cref="DocumentReader"/>) into a
/// PDF with QuestPDF. Unlike approach A's single continuous sheet, QuestPDF paginates the flow
/// automatically.
///
/// QuestPDF is dual-licensed: the source is MIT, but a paid license is required above $1M annual
/// revenue. The license is acknowledged once at startup (see Program.cs). This class assumes it
/// has been set.
///
/// Known fidelity gap: multi-column (two-column) sections have no native QuestPDF primitive, so
/// they render single-column. Documented in the README.
/// </summary>
public sealed class QuestPdfConverter : IPdfConverter
{
    public string Engine => "oss";

    // The model is in CSS px @96dpi; PDF/QuestPDF works in points (1/72"). px → pt = px * 72/96.
    private static float Pt(double px) => (float)(px * 72.0 / 96.0);

    public byte[] Convert(string docxPath)
    {
        using var stream = File.OpenRead(docxPath);
        return Convert(stream);
    }

    public byte[] Convert(Stream stream)
    {
        var model = DocumentReader.Read(stream);
        var page = model.Sections.Count > 0 ? model.Sections[0].Page : null;

        return Document.Create(doc =>
            {
                doc.Page(p =>
                {
                    if (page is not null)
                    {
                        p.Size(Pt(page.WidthPx), Pt(page.HeightPx), Unit.Point);
                        p.MarginTop(Pt(page.MarginTopPx));
                        p.MarginRight(Pt(page.MarginRightPx));
                        p.MarginBottom(Pt(page.MarginBottomPx));
                        p.MarginLeft(Pt(page.MarginLeftPx));
                    }
                    else
                    {
                        p.Margin(72f / 2f); // 0.5"
                    }

                    p.DefaultTextStyle(t =>
                        t.FontFamily("Calibri", "Carlito", "sans-serif")
                            .FontSize(11));

                    p.Content().Column(col =>
                    {
                        foreach (var section in model.Sections)
                        {
                            foreach (var block in section.Blocks)
                            {
                                col.Item().Element(c => RenderBlock(c, block));
                            }
                        }
                    });
                });
            })
            .GeneratePdf();
    }
    private static void RenderBlock(IContainer c, Block block)
    {
        switch (block)
        {
            case ParagraphBlock para:
                RenderParagraph(c, para);
                break;
            case TableBlock table:
                RenderTable(c, table);
                break;
        }
    }

    private static void RenderParagraph(IContainer c, ParagraphBlock p)
    {
        // A paragraph that is nothing but a page break.
        if (p.Runs.Count > 0 && p.Runs.All(r => r.Break == "page"))
        {
            c.PageBreak();
            return;
        }

        // Vertical spacing + left indent (list indent or paragraph indent).
        if (p.SpacingBeforePx is > 0) c = c.PaddingTop(Pt(p.SpacingBeforePx.Value));
        if (p.SpacingAfterPx is > 0) c = c.PaddingBottom(Pt(p.SpacingAfterPx.Value));
        if (p.IndentLeftPx is > 0) c = c.PaddingLeft(Pt(p.IndentLeftPx.Value));

        var textRuns = p.Runs.Where(r => r.Image is null && r.Break is null && r.Text is not null).ToList();
        var imageRuns = p.Runs.Where(r => r.Image is not null).Select(r => r.Image!).ToList();

        // Default heading appearance when the resolved runs don't already carry size/weight.
        double? headingSize = p.HeadingLevel switch { 1 => 20, 2 => 16, 3 => 13, _ => null };

        c.Column(col =>
        {
            if (textRuns.Count > 0 || p.List is not null)
            {
                col.Item().Text(text =>
                {
                    switch (p.Align)
                    {
                        case "center": text.AlignCenter(); break;
                        case "right": text.AlignRight(); break;
                        default: text.AlignLeft(); break; // "justify" falls back to left
                    }

                    if (p.List is not null)
                        text.Span($"{p.List.Marker}  ");

                    foreach (var run in textRuns)
                        ApplyRun(text.Span(run.Text!), run, p.HeadingLevel, headingSize);
                });
            }

            foreach (var img in imageRuns)
            {
                var bytes = DataUrlBytes(img.Src);
                if (bytes is null) continue;
                col.Item().MaxWidth(Pt(img.WidthPx)).Image(bytes);
            }
        });
    }

    private static void ApplyRun(TextSpanDescriptor span, Run run, int? headingLevel, double? headingSize)
    {
        var size = run.SizePt ?? headingSize;
        if (size is not null) span.FontSize((float)size.Value);

        if (run.Bold == true || (headingLevel is not null && run.Bold is null)) span.Bold();
        if (run.Italic == true) span.Italic();
        if (run.Underline == true) span.Underline();
        if (run.Strike == true) span.Strikethrough();
        if (!string.IsNullOrEmpty(run.Font)) span.FontFamily(run.Font);
        if (IsHex(run.Color)) span.FontColor(run.Color!);
        if (IsHex(run.Highlight)) span.BackgroundColor(run.Highlight!);
    }

    private static void RenderTable(IContainer c, TableBlock t)
    {
        // Number of grid columns: explicit widths if present, else the widest row's span sum.
        int columns = t.ColWidthsPx.Count > 0
            ? t.ColWidthsPx.Count
            : t.Rows.Count == 0 ? 0 : t.Rows.Max(r => r.Cells.Sum(cell => cell.GridSpan));
        if (columns == 0) return;

        // Precompute each cell's starting grid column (from gridSpan) and its rowspan (from vMerge),
        // mirroring DocView.tableLayout so colspan/rowspan match the HTML renderer.
        var starts = t.Rows.Select(row =>
        {
            int g = 0;
            return row.Cells.Select(cell => { int s = g; g += cell.GridSpan; return s; }).ToArray();
        }).ToArray();

        c.Table(table =>
        {
            // Relative (proportional) columns keep the table within the page width regardless of the
            // source grid widths — constant widths can sum wider than the printable area and throw.
            table.ColumnsDefinition(cols =>
            {
                if (t.ColWidthsPx.Count > 0)
                    foreach (var w in t.ColWidthsPx) cols.RelativeColumn((float)Math.Max(w, 1));
                else
                    for (int i = 0; i < columns; i++) cols.RelativeColumn();
            });

            for (int ri = 0; ri < t.Rows.Count; ri++)
            {
                var row = t.Rows[ri];
                for (int ci = 0; ci < row.Cells.Count; ci++)
                {
                    var cell = row.Cells[ci];
                    if (cell.VMerge == "continue") continue; // covered by the "restart" cell above

                    int start = starts[ri][ci];
                    int rowSpan = 1;
                    if (cell.VMerge == "restart")
                    {
                        for (int rj = ri + 1; rj < t.Rows.Count; rj++)
                        {
                            bool extends = false;
                            for (int k = 0; k < starts[rj].Length; k++)
                            {
                                if (starts[rj][k] == start && t.Rows[rj].Cells[k].VMerge == "continue")
                                {
                                    extends = true;
                                    break;
                                }
                            }
                            if (extends) rowSpan++;
                            else break;
                        }
                    }

                    var td = table.Cell().Row((uint)(ri + 1)).Column((uint)(start + 1));
                    if (cell.GridSpan > 1) td = td.ColumnSpan((uint)cell.GridSpan);
                    if (rowSpan > 1) td = td.RowSpan((uint)rowSpan);

                    var container = td
                        .Border(0.5f).BorderColor("#808080")
                        .Background(IsHex(cell.Shading) ? cell.Shading!
                            : row.IsHeader ? "#F2F2F2" : "#FFFFFF")
                        .Padding(3);
                    if (row.IsHeader) container = container.DefaultTextStyle(x => x.SemiBold());

                    container.Column(inner =>
                    {
                        foreach (var b in cell.Blocks)
                            inner.Item().Element(x => RenderBlock(x, b));
                    });
                }
            }
        });
    }

    private static bool IsHex(string? s) => s is not null && s.StartsWith('#');

    /// <summary>Extract raw bytes from a <c>data:*;base64,…</c> URL; null if not a data URL.</summary>
    private static byte[]? DataUrlBytes(string src)
    {
        int comma = src.IndexOf(',');
        if (comma < 0 || !src.StartsWith("data:", StringComparison.Ordinal)) return null;
        try { return System.Convert.FromBase64String(src[(comma + 1)..]); }
        catch (FormatException) { return null; }
    }
}
