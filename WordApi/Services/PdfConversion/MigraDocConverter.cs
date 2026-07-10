using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MigraDoc.Rendering;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Md = MigraDoc.DocumentObjectModel;
using MdShapes = MigraDoc.DocumentObjectModel.Shapes;
using MdTables = MigraDoc.DocumentObjectModel.Tables;

namespace WordApi.Services.PdfConversion;

/// <summary>
/// Free (MIT) low-cost engine. Unlike <see cref="QuestPdfConverter"/> — which rides the shared
/// <c>DocumentReader</c> / <c>DocumentModel</c> pipeline (approach A) — this converter is a
/// <b>self-contained</b>, independent path: it opens the <c>.docx</c> itself with the Open XML SDK
/// and walks the body directly into MigraDoc's document object model, then renders with PDFsharp.
/// Nothing here is shared with the QuestPDF implementation, so the two engines exercise fully
/// independent read-and-render pipelines (a fairer whole-pipeline comparison).
///
/// Scope: reproduce <c>word-demo.docx</c> (the same feature set the app cares about) — headings,
/// ordered/bulleted lists, run formatting, tables with grid/vertical merges, and inline images.
/// It is deliberately not a general-purpose docx engine.
///
/// PDFsharp/MigraDoc is MIT-licensed and free even for commercial use, so — unlike QuestPDF — no
/// license acknowledgement is required.
///
/// Known gaps (documented in the README): a two-column section renders single-column (MigraDoc has
/// no native multi-column primitive, same as QuestPDF); MigraDoc's DOM cannot express strikethrough
/// or text-highlight runs, so those degrade to plain text.
/// </summary>
public sealed class MigraDocConverter : IPdfConverter
{
    public string Engine => "migradoc";

    static MigraDocConverter()
    {
        // PDFsharp — unlike QuestPDF's bundled Skia/HarfBuzz — ships no fonts or text shaper; it
        // resolves glyphs through an IFontResolver we must provide, or rendering throws. (Its
        // UseWindowsFontsUnderWindows fallback only covers a few standard families, not Calibri.)
        // FileSystemFontResolver reads .ttf/.otf from the OS font directories, so this works on
        // Windows now and on Linux once Carlito/Liberation are installed — a real deployment cost
        // the free engine carries that QuestPDF does not. See the README's deploy notes.
        PdfSharp.Fonts.GlobalFontSettings.FontResolver = FileSystemFontResolver.Instance;
    }

    // Word measures in twips (1/20 pt) and EMU; MigraDoc/PDFsharp work in points (1/72").
    private const double TwipsToPt = 72.0 / 1440.0;    // 1440 twips = 1 inch = 72 pt
    private const double EmuToPt = 72.0 / 914400.0;    // 914400 EMU = 1 inch = 72 pt

    
    public byte[] Convert(string docxPath)
    {
        using var stream = File.OpenRead(docxPath);
        return Convert(stream);
    }
    
    
    public byte[] Convert(Stream stream)
    {
        // Read-only, ReadWrite share so a copy open in Word doesn't block us (same as DocumentReader).
       
        using var wdoc = WordprocessingDocument.Open(stream, false);
        var main = wdoc.MainDocumentPart
            ?? throw new InvalidOperationException("Document has no main part.");
        var body = main.Document?.Body
            ?? throw new InvalidOperationException("Document has no body.");

        var styles = new StyleIndex(main);
        var numbering = new NumberingIndex(main);
        var counters = new Dictionary<(int, int), int>();

        var doc = new Md.Document();
        var section = doc.AddSection();
        double contentWidthPt = ApplyPageSetup(section, body);

        // Default font for the whole document; single family (MigraDoc has no fallback list).
        // On Linux install/register the metric-compatible "Carlito" — see the README.
        var normal = doc.Styles["Normal"]!; // MigraDoc always defines the built-in "Normal" style
        normal.Font.Name = "Calibri";
        normal.Font.Size = Md.Unit.FromPoint(11);

        bool pageBreakPending = false;
        foreach (var el in body.ChildElements)
        {
            try
            {
                switch (el)
                {
                    case Paragraph p:
                        // A paragraph that is only a page break defers the break to the next block.
                        if (IsPageBreakOnly(p)) { pageBreakPending = true; break; }
                        RenderParagraph(() => section.AddParagraph(), p, styles, numbering, counters,
                                        main, contentWidthPt, applyPageBreak: pageBreakPending);
                        pageBreakPending = false;
                        break;

                    case Table t:
                        if (pageBreakPending)
                        {
                            section.AddParagraph().Format.PageBreakBefore = true;
                            pageBreakPending = false;
                        }
                        RenderTable(section, t, styles, numbering, counters, main, contentWidthPt);
                        break;
                }
            }
            catch
            {
                // One unreadable block must never break the whole conversion.
            }
        }

        var renderer = new PdfDocumentRenderer { Document = doc };
        renderer.RenderDocument();
        using var outMs = new MemoryStream();
        renderer.PdfDocument.Save(outMs);
        return outMs.ToArray();
    }

    /// <summary>Page size + margins from the body's section properties; returns content width (pt).</summary>
    private static double ApplyPageSetup(Md.Section section, Body body)
    {
        // Defaults: US Letter, 1in margins.
        double w = 12240, h = 15840, mt = 1440, mr = 1440, mb = 1440, ml = 1440;

        var sect = body.Elements<SectionProperties>().LastOrDefault();
        if (sect?.GetFirstChild<PageSize>() is { } pgSz)
        {
            if (pgSz.Width?.Value is { } pw) w = pw;
            if (pgSz.Height?.Value is { } ph) h = ph;
        }
        if (sect?.GetFirstChild<PageMargin>() is { } pgMar)
        {
            if (pgMar.Top?.Value is { } t) mt = t;
            if (pgMar.Right?.Value is { } r) mr = r;
            if (pgMar.Bottom?.Value is { } b) mb = b;
            if (pgMar.Left?.Value is { } l) ml = l;
        }

        var ps = section.PageSetup;
        ps.PageWidth = Md.Unit.FromPoint(w * TwipsToPt);
        ps.PageHeight = Md.Unit.FromPoint(h * TwipsToPt);
        ps.TopMargin = Md.Unit.FromPoint(mt * TwipsToPt);
        ps.RightMargin = Md.Unit.FromPoint(mr * TwipsToPt);
        ps.BottomMargin = Md.Unit.FromPoint(mb * TwipsToPt);
        ps.LeftMargin = Md.Unit.FromPoint(ml * TwipsToPt);

        return (w - ml - mr) * TwipsToPt;
    }

    // ---- Paragraphs -----------------------------------------------------------------------

    private static bool IsPageBreakOnly(Paragraph p)
    {
        bool hasBreak = false, hasText = false;
        foreach (var br in p.Descendants<Break>())
            if (br.Type?.Value == BreakValues.Page) hasBreak = true;
        foreach (var tx in p.Descendants<Text>())
            if (!string.IsNullOrEmpty(tx.Text)) hasText = true;
        return hasBreak && !hasText;
    }

    private static void RenderParagraph(Func<Md.Paragraph> addParagraph, Paragraph p, StyleIndex styles,
        NumberingIndex numbering, Dictionary<(int, int), int> counters, MainDocumentPart main,
        double contentWidthPt, bool applyPageBreak)
    {
        var pPr = p.ParagraphProperties;
        var styleId = pPr?.ParagraphStyleId?.Val?.Value;
        var chain = styles.Chain(styleId);

        // Effective paragraph formatting: docDefaults -> style chain -> direct pPr.
        var pf = new ParaFmt();
        pf.Apply(styles.DocDefaultsPPr);
        foreach (var s in chain) pf.Apply(s.StyleParagraphProperties);
        pf.Apply(pPr);

        // Base run formatting inherited by every run.
        var baseRun = new RunFmt();
        baseRun.Apply(styles.DocDefaultsRPr);
        foreach (var s in chain) baseRun.Apply(s.StyleRunProperties);
        baseRun.Apply(pPr?.GetFirstChild<ParagraphMarkRunProperties>());

        int? heading = HeadingLevel(styleId, pf.OutlineLevel);
        double? headingSize = heading switch { 1 => 20, 2 => 16, 3 => 13, _ => null };

        // Numbering: direct numPr wins, else a style in the chain may carry it.
        var numPr = pPr?.NumberingProperties
                    ?? chain.Select(s => s.StyleParagraphProperties?.NumberingProperties)
                            .FirstOrDefault(n => n != null);
        string? marker = null;
        double? indentTwips = pf.IndentLeftTwips;
        if (numPr?.NumberingId?.Val != null)
        {
            int numId = numPr.NumberingId.Val.Value;
            int ilvl = numPr.NumberingLevelReference?.Val?.Value ?? 0;
            if (numbering.Resolve(numId, ilvl) is { } lvl)
            {
                if (lvl.Ordered)
                {
                    counters.TryGetValue((numId, ilvl), out int n);
                    n = n == 0 ? lvl.Start : n + 1;
                    counters[(numId, ilvl)] = n;
                    foreach (var key in counters.Keys.Where(k => k.Item1 == numId && k.Item2 > ilvl).ToList())
                        counters.Remove(key);
                    marker = $"{n}.";
                }
                else marker = "•";

                double? directLeft = ParseTwips(pPr?.Indentation?.Left?.Value);
                indentTwips = directLeft ?? lvl.IndentLeftTwips ?? pf.IndentLeftTwips;
            }
        }

        var mp = addParagraph();
        if (applyPageBreak) mp.Format.PageBreakBefore = true;

        mp.Format.Alignment = pf.Align switch
        {
            "center" => Md.ParagraphAlignment.Center,
            "right" => Md.ParagraphAlignment.Right,
            _ => Md.ParagraphAlignment.Left,   // "justify" falls back to left (matches QuestPDF path)
        };
        if (pf.SpacingBeforeTwips is > 0) mp.Format.SpaceBefore = Md.Unit.FromPoint(pf.SpacingBeforeTwips.Value * TwipsToPt);
        if (pf.SpacingAfterTwips is > 0) mp.Format.SpaceAfter = Md.Unit.FromPoint(pf.SpacingAfterTwips.Value * TwipsToPt);
        if (indentTwips is > 0) mp.Format.LeftIndent = Md.Unit.FromPoint(indentTwips.Value * TwipsToPt);

        if (marker is not null) mp.AddText($"{marker}  ");

        var images = new List<(byte[] bytes, double wPt)>();
        foreach (var inline in CollectInlines(p, baseRun, styles, main))
        {
            if (inline.Image is { } im) { images.Add(im); continue; }
            if (inline.Text is null) continue;

            var ft = mp.AddFormattedText(inline.Text);
            ApplyRun(ft, inline.Fmt, heading, headingSize);
        }

        // Inline pictures follow the text as their own paragraphs (mirrors the QuestPDF column).
        foreach (var (bytes, wPt) in images)
        {
            try
            {
                var ip = addParagraph();
                var img = ip.AddImage("base64:" + System.Convert.ToBase64String(bytes));
                img.LockAspectRatio = true;
                if (wPt > 0) img.Width = Md.Unit.FromPoint(Math.Min(wPt, contentWidthPt));
            }
            catch
            {
                // Unsupported image format (e.g. EMF/WMF) — skip rather than fail the document.
            }
        }
    }

    private static void ApplyRun(Md.FormattedText ft, RunFmt run, int? headingLevel, double? headingSize)
    {
        double? sizePt = run.SizeHalfPts is { } hp ? hp / 2.0 : headingSize;
        if (sizePt is not null) ft.Font.Size = Md.Unit.FromPoint(sizePt.Value);

        if (run.Bold == true || (headingLevel is not null && run.Bold is null)) ft.Bold = true;
        if (run.Italic == true) ft.Italic = true;
        if (run.Underline == true) ft.Underline = Md.Underline.Single;
        if (!string.IsNullOrEmpty(run.Font)) ft.Font.Name = run.Font;
        if (ParseColor(run.Color) is { } c) ft.Font.Color = c;
        // Strikethrough and text highlight have no MigraDoc DOM equivalent — intentionally dropped.
    }

    // ---- Inline collection ----------------------------------------------------------------

    private readonly record struct Inline(string? Text, RunFmt Fmt, (byte[] bytes, double wPt)? Image);

    private static IEnumerable<Inline> CollectInlines(OpenXmlElement container, RunFmt baseRun,
        StyleIndex styles, MainDocumentPart main)
    {
        foreach (var child in container.ChildElements)
        {
            switch (child)
            {
                case Hyperlink link:
                    foreach (var i in CollectInlines(link, baseRun, styles, main)) yield return i;
                    break;

                case Run r:
                    var fmt = baseRun.Clone();
                    var rPr = r.RunProperties;
                    foreach (var s in styles.Chain(rPr?.RunStyle?.Val?.Value)) fmt.Apply(s.StyleRunProperties);
                    fmt.Apply(rPr);

                    foreach (var el in r.ChildElements)
                    {
                        switch (el)
                        {
                            case Text t:
                                yield return new Inline(t.Text, fmt, null);
                                break;
                            case Drawing d when ReadImage(d, main) is { } img:
                                yield return new Inline(null, fmt, img);
                                break;
                        }
                    }
                    break;
            }
        }
    }

    private static (byte[] bytes, double wPt)? ReadImage(Drawing drawing, MainDocumentPart main)
    {
        try
        {
            var relId = drawing.Descendants<A.Blip>().FirstOrDefault()?.Embed?.Value;
            if (relId == null || main.GetPartById(relId) is not ImagePart part) return null;

            using var s = part.GetStream();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            var bytes = ms.ToArray();

            // PDFsharp decodes only PNG/JPEG/BMP; other formats (e.g. GIF, which the sample's Web
            // Access Symbol uses) render as an "Image has no valid type" placeholder. Skip them —
            // QuestPDF's Skia handles more formats; this is a documented MigraDoc gap.
            if (!IsSupportedImage(bytes)) return null;

            var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();
            double wPt = extent?.Cx != null ? extent.Cx.Value * EmuToPt : 0;
            return (bytes, wPt);
        }
        catch
        {
            return null;
        }
    }

    // ---- Tables ---------------------------------------------------------------------------

    private static void RenderTable(Md.Section section, Table t, StyleIndex styles,
        NumberingIndex numbering, Dictionary<(int, int), int> counters, MainDocumentPart main,
        double contentWidthPt)
    {
        var rows = t.Elements<TableRow>().ToList();
        if (rows.Count == 0) return;

        // Grid column widths (twips → pt); fall back to equal split of the widest row.
        var gridWidthsPt = new List<double>();
        if (t.GetFirstChild<TableGrid>() is { } grid)
            foreach (var gc in grid.Elements<GridColumn>())
                gridWidthsPt.Add((ParseTwips(gc.Width?.Value) ?? 0) * TwipsToPt);

        int columns = gridWidthsPt.Count > 0
            ? gridWidthsPt.Count
            : rows.Max(r => r.Elements<TableCell>().Sum(GridSpanOf));
        if (columns == 0) return;

        // Scale columns so the table always fits the content width (MigraDoc has no auto-fit,
        // and constant widths can sum wider than the page — QuestPDF sidesteps this with relatives).
        double[] widthsPt = new double[columns];
        double totalPt = gridWidthsPt.Take(columns).Sum();
        for (int i = 0; i < columns; i++)
        {
            double raw = i < gridWidthsPt.Count && gridWidthsPt[i] > 0 ? gridWidthsPt[i] : 0;
            widthsPt[i] = raw;
        }
        if (totalPt <= 0) { for (int i = 0; i < columns; i++) widthsPt[i] = contentWidthPt / columns; totalPt = contentWidthPt; }
        double scale = totalPt > contentWidthPt ? contentWidthPt / totalPt : 1.0;

        var table = section.AddTable();
        table.Borders.Width = Md.Unit.FromPoint(0.5);
        table.Borders.Color = new Md.Color(0x80, 0x80, 0x80);
        for (int i = 0; i < columns; i++)
            table.AddColumn(Md.Unit.FromPoint(Math.Max(widthsPt[i] * scale, 1)));

        // Precompute each cell's starting grid column (from gridSpan), mirroring the QuestPDF path.
        var cellsPerRow = rows.Select(r => r.Elements<TableCell>().ToList()).ToList();
        var starts = cellsPerRow.Select(cells =>
        {
            int g = 0;
            return cells.Select(c => { int s = g; g += GridSpanOf(c); return s; }).ToArray();
        }).ToList();

        for (int ri = 0; ri < rows.Count; ri++)
        {
            var mRow = table.AddRow();
            var cells = cellsPerRow[ri];
            for (int ci = 0; ci < cells.Count; ci++)
            {
                var cell = cells[ci];
                var vMerge = VMergeOf(cell);
                if (vMerge == "continue") continue; // covered by the "restart" cell above

                int start = starts[ri][ci];
                if (start >= columns) continue;
                int span = GridSpanOf(cell);

                int rowSpan = 1;
                if (vMerge == "restart")
                {
                    for (int rj = ri + 1; rj < rows.Count; rj++)
                    {
                        bool extends = false;
                        for (int k = 0; k < starts[rj].Length; k++)
                            if (starts[rj][k] == start && VMergeOf(cellsPerRow[rj][k]) == "continue")
                            {
                                extends = true;
                                break;
                            }
                        if (extends) rowSpan++;
                        else break;
                    }
                }

                var mCell = mRow.Cells[start];
                if (span > 1) mCell.MergeRight = Math.Min(span - 1, columns - start - 1);
                if (rowSpan > 1) mCell.MergeDown = rowSpan - 1;

                var shading = ShadingOf(cell);
                if (ParseColor(shading) is { } fill) mCell.Shading.Color = fill;
                else if (ri == 0) mCell.Shading.Color = new Md.Color(0xF2, 0xF2, 0xF2);
                if (ri == 0) mCell.Format.Font.Bold = true; // header row

                foreach (var block in cell.ChildElements)
                    if (block is Paragraph cp)
                        RenderParagraph(() => mCell.AddParagraph(), cp, styles, numbering, counters,
                                        main, contentWidthPt, applyPageBreak: false);
                // Nested tables inside cells are unsupported by MigraDoc; the sample has none.
            }
        }
    }

    /// <summary>True for PNG / JPEG / BMP magic bytes — the formats PDFsharp can decode.</summary>
    private static bool IsSupportedImage(byte[] b) =>
        (b.Length >= 4 && b[0] == 0x89 && b[1] == 0x50 && b[2] == 0x4E && b[3] == 0x47)   // PNG
        || (b.Length >= 3 && b[0] == 0xFF && b[1] == 0xD8 && b[2] == 0xFF)                 // JPEG
        || (b.Length >= 2 && b[0] == 0x42 && b[1] == 0x4D);                                // BMP

    private static int GridSpanOf(TableCell c) => c.TableCellProperties?.GridSpan?.Val?.Value ?? 1;

    private static string? VMergeOf(TableCell c)
    {
        var vm = c.TableCellProperties?.VerticalMerge;
        if (vm == null) return null;
        return vm.Val?.Value == MergedCellValues.Restart ? "restart" : "continue";
    }

    private static string? ShadingOf(TableCell c)
    {
        var fill = c.TableCellProperties?.Shading?.Fill?.Value;
        return fill != null && fill != "auto" ? "#" + fill : null;
    }

    // ---- Helpers --------------------------------------------------------------------------

    private static int? HeadingLevel(string? styleId, int? outlineLevel)
    {
        if (styleId != null)
        {
            var m = Regex.Match(styleId, @"^Heading(\d)$", RegexOptions.IgnoreCase);
            if (m.Success) return int.Parse(m.Groups[1].Value);
        }
        return outlineLevel is { } o ? o + 1 : null;
    }

    private static double? ParseTwips(string? s) =>
        s != null && double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? v : null;

    /// <summary>Parse a <c>#RRGGBB</c> hex string into a MigraDoc color; null if not valid hex.</summary>
    private static Md.Color? ParseColor(string? hex)
    {
        if (hex is null || hex.Length != 7 || hex[0] != '#') return null;
        return byte.TryParse(hex.AsSpan(1, 2), NumberStyles.HexNumber, null, out var r)
            && byte.TryParse(hex.AsSpan(3, 2), NumberStyles.HexNumber, null, out var g)
            && byte.TryParse(hex.AsSpan(5, 2), NumberStyles.HexNumber, null, out var b)
                ? new Md.Color(r, g, b)
                : null;
    }

    // ---- Effective-formatting accumulators (independent of DocumentReader) ----------------

    private sealed class RunFmt
    {
        public bool? Bold { get; private set; }
        public bool? Italic { get; private set; }
        public bool? Underline { get; private set; }
        public string? Font { get; private set; }
        public double? SizeHalfPts { get; private set; }
        public string? Color { get; private set; }

        public RunFmt Clone() => new()
        {
            Bold = Bold, Italic = Italic, Underline = Underline,
            Font = Font, SizeHalfPts = SizeHalfPts, Color = Color,
        };

        public void Apply(OpenXmlElement? props)
        {
            if (props == null) return;
            if (props.GetFirstChild<Bold>() is { } b) Bold = b.Val?.Value ?? true;
            if (props.GetFirstChild<Italic>() is { } i) Italic = i.Val?.Value ?? true;
            if (props.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Underline>() is { } u)
                Underline = (u.Val?.Value ?? UnderlineValues.Single) != UnderlineValues.None;
            if (props.GetFirstChild<RunFonts>()?.Ascii?.Value is { } f) Font = f;
            if (props.GetFirstChild<FontSize>()?.Val?.Value is { } sz && ParseTwips(sz) is { } hp) SizeHalfPts = hp;
            if (props.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Color>()?.Val?.Value is { } c && c != "auto")
                Color = "#" + c;
        }
    }

    private sealed class ParaFmt
    {
        public string? Align { get; private set; }
        public double? IndentLeftTwips { get; private set; }
        public double? SpacingBeforeTwips { get; private set; }
        public double? SpacingAfterTwips { get; private set; }
        public int? OutlineLevel { get; private set; }

        public void Apply(OpenXmlElement? pPr)
        {
            if (pPr == null) return;
            if (pPr.GetFirstChild<Justification>()?.Val?.Value is { } jc)
                Align = jc == JustificationValues.Both ? "justify" : jc.ToString()?.ToLowerInvariant();
            if (ParseTwips(pPr.GetFirstChild<Indentation>()?.Left?.Value) is { } left) IndentLeftTwips = left;
            var sp = pPr.GetFirstChild<SpacingBetweenLines>();
            if (ParseTwips(sp?.Before?.Value) is { } bf) SpacingBeforeTwips = bf;
            if (ParseTwips(sp?.After?.Value) is { } af) SpacingAfterTwips = af;
            if (pPr.GetFirstChild<OutlineLevel>()?.Val?.Value is { } ol) OutlineLevel = ol;
        }
    }

    // ---- Style + numbering indices (independent, trimmed to what the sample needs) ---------

    private sealed class StyleIndex
    {
        private readonly Dictionary<string, Style> _byId = [];
        public OpenXmlElement? DocDefaultsRPr { get; }
        public OpenXmlElement? DocDefaultsPPr { get; }

        public StyleIndex(MainDocumentPart main)
        {
            var styles = main.StyleDefinitionsPart?.Styles;
            if (styles == null) return;
            foreach (var s in styles.Elements<Style>())
                if (s.StyleId?.Value is { } id) _byId[id] = s;
            DocDefaultsRPr = styles.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
            DocDefaultsPPr = styles.DocDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle;
        }

        /// <summary>Chain from the root base style down to <paramref name="styleId"/> (inclusive).</summary>
        public List<Style> Chain(string? styleId)
        {
            var chain = new List<Style>();
            var seen = new HashSet<string>();
            var id = styleId;
            while (id != null && seen.Add(id) && _byId.TryGetValue(id, out var s))
            {
                chain.Add(s);
                id = s.BasedOn?.Val?.Value;
            }
            chain.Reverse(); // base first, most-derived last
            return chain;
        }
    }

    private sealed record LevelInfo(bool Ordered, int Start, double? IndentLeftTwips);

    private sealed class NumberingIndex
    {
        private readonly Dictionary<int, int> _numToAbstract = [];
        private readonly Dictionary<int, AbstractNum> _abstractById = [];

        public NumberingIndex(MainDocumentPart main)
        {
            var num = main.NumberingDefinitionsPart?.Numbering;
            if (num == null) return;
            foreach (var an in num.Elements<AbstractNum>())
                if (an.AbstractNumberId?.Value is { } aid) _abstractById[aid] = an;
            foreach (var n in num.Elements<NumberingInstance>())
                if (n.NumberID?.Value is { } id && n.AbstractNumId?.Val?.Value is { } aid)
                    _numToAbstract[id] = aid;
        }

        public LevelInfo? Resolve(int numId, int ilvl)
        {
            if (!_numToAbstract.TryGetValue(numId, out int aid)) return null;
            if (!_abstractById.TryGetValue(aid, out var an)) return null;
            var lvl = an.Elements<Level>().FirstOrDefault(l => (l.LevelIndex?.Value ?? 0) == ilvl);
            if (lvl == null) return null;

            var fmt = lvl.NumberingFormat?.Val?.Value;
            bool ordered = fmt != null && fmt != NumberFormatValues.Bullet && fmt != NumberFormatValues.None;
            int start = lvl.StartNumberingValue?.Val?.Value ?? 1;
            double? indLeft = ParseTwips(lvl.PreviousParagraphProperties?.Indentation?.Left?.Value);
            return new LevelInfo(ordered, start, indLeft);
        }
    }
}

/// <summary>
/// Minimal cross-platform <see cref="PdfSharp.Fonts.IFontResolver"/> for the MigraDoc engine.
/// PDFsharp ships no fonts, so we map family + bold/italic to a .ttf/.otf file found in the OS font
/// directories (Windows Fonts, the standard Linux font paths, and an app-relative <c>fonts/</c>
/// folder). Unknown families fall back to Calibri → Carlito → Arial. This is what makes the free
/// engine deployable on Linux: install Carlito/Liberation and this resolver picks them up.
/// </summary>
internal sealed class FileSystemFontResolver : PdfSharp.Fonts.IFontResolver
{
    public static readonly FileSystemFontResolver Instance = new();

    // Family → (regular, bold, italic, bold-italic) file names, in OS + open-source variants.
    private static readonly Dictionary<string, string[]> Families = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Calibri"] = ["calibri.ttf", "calibrib.ttf", "calibrii.ttf", "calibriz.ttf"],
        ["Carlito"] = ["Carlito-Regular.ttf", "Carlito-Bold.ttf", "Carlito-Italic.ttf", "Carlito-BoldItalic.ttf"],
        ["Arial"] = ["arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf"],
        ["Liberation Sans"] = ["LiberationSans-Regular.ttf", "LiberationSans-Bold.ttf", "LiberationSans-Italic.ttf", "LiberationSans-BoldItalic.ttf"],
        ["Times New Roman"] = ["times.ttf", "timesbd.ttf", "timesi.ttf", "timesbi.ttf"],
        ["Courier New"] = ["cour.ttf", "courbd.ttf", "couri.ttf", "courbi.ttf"],
    };

    // filename (lower) → absolute path, built once by scanning the font directories.
    private static readonly Dictionary<string, string> FileIndex = BuildFileIndex();
    private readonly Dictionary<string, byte[]?> _fontCache = new(StringComparer.OrdinalIgnoreCase);
    private readonly System.Threading.Lock _lock = new();

    public PdfSharp.Fonts.FontResolverInfo? ResolveTypeface(string familyName, bool bold, bool italic)
        => new($"{familyName}|{(bold ? "b" : "")}{(italic ? "i" : "")}");

    public byte[]? GetFont(string faceName)
    {
        lock (_lock)
        {
            if (_fontCache.TryGetValue(faceName, out var cached)) return cached;

            var parts = faceName.Split('|');
            string family = parts[0];
            bool bold = parts.Length > 1 && parts[1].Contains('b');
            bool italic = parts.Length > 1 && parts[1].Contains('i');

            var bytes = Load(family, bold, italic)
                        ?? Load("Calibri", bold, italic)
                        ?? Load("Carlito", bold, italic)
                        ?? Load("Arial", bold, italic)
                        ?? Load("Arial", false, false);
            return _fontCache[faceName] = bytes;
        }
    }

    private static byte[]? Load(string family, bool bold, bool italic)
    {
        if (!Families.TryGetValue(family, out var files)) return null;
        int idx = (bold, italic) switch { (true, true) => 3, (true, false) => 1, (false, true) => 2, _ => 0 };
        // Prefer the exact style; fall back to the regular face if that variant isn't installed.
        foreach (var candidate in new[] { files[idx], files[0] })
            if (FileIndex.TryGetValue(candidate, out var path))
                try { return File.ReadAllBytes(path); } catch { /* unreadable — try next */ }
        return null;
    }

    private static Dictionary<string, string> BuildFileIndex()
    {
        var dirs = new List<string>
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts"),
            "/usr/share/fonts", "/usr/local/share/fonts",
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".fonts"),
            Path.Combine(AppContext.BaseDirectory, "fonts"),
        };

        var index = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var dir in dirs)
        {
            if (string.IsNullOrEmpty(dir) || !Directory.Exists(dir)) continue;
            try
            {
                foreach (var pattern in new[] { "*.ttf", "*.otf" })
                    foreach (var file in Directory.EnumerateFiles(dir, pattern, SearchOption.AllDirectories))
                        index.TryAdd(Path.GetFileName(file), file);
            }
            catch { /* skip unreadable font directory */ }
        }
        return index;
    }
}
