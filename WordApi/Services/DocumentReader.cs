using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordApi.Models;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using MRun = WordApi.Models.Run;   // model run
using WRun = DocumentFormat.OpenXml.Wordprocessing.Run;             // OpenXML run
using WTable = DocumentFormat.OpenXml.Wordprocessing.Table;         // OpenXML table
using WTableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;   // OpenXML row
using WTableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell; // OpenXML cell
using MTableRow = WordApi.Models.TableRow;                          // model row
using MTableCell = WordApi.Models.TableCell;                        // model cell

namespace WordApi.Services;

/// <summary>
/// Reads a .docx into the flat <see cref="DocumentModel"/> (approach A). Style inheritance,
/// numbering, and images are all resolved here so the browser gets compact, effective JSON —
/// the same "backend resolves, frontend just draws" split as the Excel POC's ClosedXML reader.
///
/// Read-only: the file is opened <c>false</c> (no edit) with <c>FileShare.ReadWrite</c> so a
/// copy open in Word doesn't block. Every item is wrapped so one unreadable element can't
/// break the whole preview.
/// </summary>
public static class DocumentReader
{
    // Unit conversions. 1 inch = 1440 twips (dxa) = 914400 EMU = 96 CSS px.
    private const double TwipsToPx = 96.0 / 1440.0;   // == 1/15
    private const double EmuToPx = 96.0 / 914400.0;    // == 1/9525

    public static DocumentModel Read(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var doc = WordprocessingDocument.Open(fs, false);
        var main = doc.MainDocumentPart
            ?? throw new InvalidOperationException("Document has no main part.");
        var body = main.Document.Body
            ?? throw new InvalidOperationException("Document has no body.");

        var ctx = new Ctx(main, new StyleIndex(main), new NumberingIndex(main));

        var sections = new List<Section>();
        var current = new List<Block>();

        foreach (var el in body.ChildElements)
        {
            try
            {
                switch (el)
                {
                    case Paragraph p:
                        current.Add(ctx.BuildParagraph(p));
                        // A sectPr inside a paragraph ends the current section.
                        var inlineSect = p.ParagraphProperties?.GetFirstChild<SectionProperties>();
                        if (inlineSect != null)
                        {
                            sections.Add(BuildSection(inlineSect, current));
                            current = [];
                        }
                        break;

                    case WTable t:
                        current.Add(ctx.BuildTable(t));
                        break;

                    // The body's final child is the last section's properties.
                    case SectionProperties bodySect:
                        sections.Add(BuildSection(bodySect, current));
                        current = [];
                        break;
                }
            }
            catch
            {
                // A single unreadable block must never break the preview.
            }
        }

        if (current.Count > 0)
            sections.Add(BuildSection(null, current));

        return new DocumentModel(sections, ctx.Stats());
    }

    private static Section BuildSection(SectionProperties? sect, List<Block> blocks)
    {
        // Defaults: US Letter, 1in margins, single column.
        double w = 12240, h = 15840, mt = 1440, mr = 1440, mb = 1440, ml = 1440;
        int cols = 1;

        var pgSz = sect?.GetFirstChild<PageSize>();
        if (pgSz != null)
        {
            if (pgSz.Width != null) w = pgSz.Width.Value;
            if (pgSz.Height != null) h = pgSz.Height.Value;
        }
        var pgMar = sect?.GetFirstChild<PageMargin>();
        if (pgMar != null)
        {
            if (pgMar.Top != null) mt = pgMar.Top.Value;
            if (pgMar.Right != null) mr = pgMar.Right.Value;
            if (pgMar.Bottom != null) mb = pgMar.Bottom.Value;
            if (pgMar.Left != null) ml = pgMar.Left.Value;
        }
        var colsEl = sect?.GetFirstChild<Columns>();
        if (colsEl?.ColumnCount != null) cols = colsEl.ColumnCount.Value;

        var page = new PageInfo(
            w * TwipsToPx, h * TwipsToPx,
            mt * TwipsToPx, mr * TwipsToPx, mb * TwipsToPx, ml * TwipsToPx);
        return new Section(cols, page, blocks);
    }

    // ---- Walk context: holds the resolvers + running counters/stats -----------------------

    private sealed class Ctx(MainDocumentPart main, StyleIndex styles, NumberingIndex numbering)
    {
        // Ordered-list counters, keyed per (numId, level) so interleaved lists keep separate runs.
        private readonly Dictionary<(int, int), int> _counters = [];
        private int _paras, _words, _headings, _tables, _images;

        public DocumentStats Stats() => new(_paras, _words, _headings, _tables, _images);

        public ParagraphBlock BuildParagraph(Paragraph p)
        {
            _paras++;
            var pPr = p.ParagraphProperties;
            var styleId = pPr?.ParagraphStyleId?.Val?.Value;
            var chain = styles.Chain(styleId);

            // Effective paragraph-level formatting: docDefaults -> style chain -> direct pPr.
            var pf = new ParaFmt();
            pf.Apply(styles.DocDefaultsPPr);
            foreach (var s in chain) pf.Apply(s.StyleParagraphProperties);
            pf.Apply(pPr);

            // Base run formatting inherited by every run in the paragraph.
            var baseRun = new RunFmt();
            baseRun.Apply(styles.DocDefaultsRPr);
            foreach (var s in chain) baseRun.Apply(s.StyleRunProperties);
            baseRun.Apply(pPr?.GetFirstChild<ParagraphMarkRunProperties>());

            int? heading = HeadingLevel(styleId, pf.OutlineLevel);
            if (heading != null) _headings++;

            // Numbering: direct numPr wins, else a style in the chain may carry it.
            var numPr = pPr?.NumberingProperties
                        ?? chain.Select(s => s.StyleParagraphProperties?.NumberingProperties)
                                .FirstOrDefault(n => n != null);
            ListInfo? list = null;
            double? indentPx = pf.IndentLeftTwips is { } dl ? dl * TwipsToPx : null;
            if (numPr?.NumberingId?.Val != null)
            {
                int numId = numPr.NumberingId.Val.Value;
                int ilvl = numPr.NumberingLevelReference?.Val?.Value ?? 0;
                var lvl = numbering.Resolve(numId, ilvl);
                if (lvl != null)
                {
                    string marker;
                    if (lvl.Ordered)
                    {
                        _counters.TryGetValue((numId, ilvl), out int n);
                        n = n == 0 ? lvl.Start : n + 1;
                        _counters[(numId, ilvl)] = n;
                        // Restart any deeper levels of this list.
                        foreach (var key in _counters.Keys.Where(k => k.Item1 == numId && k.Item2 > ilvl).ToList())
                            _counters.Remove(key);
                        marker = $"{n}.";
                    }
                    else marker = "•"; // bullet
                    list = new ListInfo(lvl.Ordered, ilvl, marker);
                    // For list items the numbering level's indent (which distinguishes the
                    // bulleted sub-list at 1080 twips from the numbered list at 720) is more
                    // specific than the ListParagraph style's flat indent. Precedence:
                    // direct paragraph indent > numbering-level indent > style indent.
                    double? directLeft = null;
                    if (pPr?.Indentation?.Left?.Value is { } dlv
                        && double.TryParse(dlv, NumberStyles.Any, CultureInfo.InvariantCulture, out var dpv))
                        directLeft = dpv;
                    var chosen = directLeft ?? lvl.IndentLeftTwips ?? pf.IndentLeftTwips;
                    if (chosen is { } cv) indentPx = cv * TwipsToPx;
                }
            }

            var runs = new List<MRun>();
            string? anchor = null;
            CollectRuns(p, baseRun, null, runs, ref anchor);

            _words += CountWords(runs);

            return new ParagraphBlock(
                runs,
                StyleId: styleId,
                HeadingLevel: heading,
                List: list,
                Align: pf.Align,
                IndentLeftPx: indentPx,
                SpacingBeforePx: pf.SpacingBeforeTwips is { } sb ? sb * TwipsToPx : null,
                SpacingAfterPx: pf.SpacingAfterTwips is { } sa ? sa * TwipsToPx : null,
                Anchor: anchor);
        }

        /// <summary>Walk a paragraph's inline content (runs, hyperlinks, bookmarks) in order.</summary>
        private void CollectRuns(OpenXmlElement container, RunFmt baseRun, string? href,
                                 List<MRun> runs, ref string? anchor)
        {
            foreach (var child in container.ChildElements)
            {
                switch (child)
                {
                    case BookmarkStart bm when bm.Name != null:
                        anchor ??= bm.Name.Value;
                        break;

                    case Hyperlink link:
                        string? target = link.Anchor != null
                            ? "#" + link.Anchor.Value
                            : ResolveHyperlink(link.Id?.Value);
                        CollectRuns(link, baseRun, target, runs, ref anchor);
                        break;

                    case WRun r:
                        AppendRun(r, baseRun, href, runs);
                        break;
                }
            }
        }

        private void AppendRun(WRun r, RunFmt baseRun, string? href, List<MRun> runs)
        {
            // Effective run formatting: paragraph base -> char-style chain (rStyle) -> direct rPr.
            var fmt = baseRun.Clone();
            var rPr = r.RunProperties;
            var rStyle = rPr?.RunStyle?.Val?.Value;
            foreach (var s in styles.Chain(rStyle)) fmt.Apply(s.StyleRunProperties);
            fmt.Apply(rPr);

            foreach (var el in r.ChildElements)
            {
                switch (el)
                {
                    case Text t:
                        runs.Add(fmt.ToRun(t.Text, href));
                        break;

                    case Break br:
                        var kind = br.Type?.Value switch
                        {
                            var v when v == BreakValues.Page => "page",
                            var v when v == BreakValues.Column => "column",
                            _ => "line",
                        };
                        runs.Add(new MRun(Break: kind));
                        break;

                    case Drawing d:
                        var img = ReadImage(d);
                        if (img != null) { runs.Add(new MRun(Image: img)); _images++; }
                        break;
                }
            }
        }

        private InlineImage? ReadImage(Drawing drawing)
        {
            try
            {
                var blip = drawing.Descendants<A.Blip>().FirstOrDefault();
                var relId = blip?.Embed?.Value;
                if (relId == null) return null;
                if (main.GetPartById(relId) is not ImagePart part) return null;

                byte[] bytes;
                using (var s = part.GetStream())
                using (var ms = new MemoryStream())
                {
                    s.CopyTo(ms);
                    bytes = ms.ToArray();
                }
                string src = $"data:{part.ContentType};base64,{Convert.ToBase64String(bytes)}";

                var extent = drawing.Descendants<DW.Extent>().FirstOrDefault();
                double wpx = extent?.Cx != null ? extent.Cx.Value * EmuToPx : 0;
                double hpx = extent?.Cy != null ? extent.Cy.Value * EmuToPx : 0;

                var docPr = drawing.Descendants<DW.DocProperties>().FirstOrDefault();
                string? alt = docPr?.Description?.Value ?? docPr?.Title?.Value;

                return new InlineImage(src, wpx, hpx, alt);
            }
            catch
            {
                return null;
            }
        }

        private string? ResolveHyperlink(string? relId)
        {
            if (relId == null) return null;
            try
            {
                var rel = main.HyperlinkRelationships.FirstOrDefault(h => h.Id == relId);
                return rel?.Uri?.ToString();
            }
            catch { return null; }
        }

        public TableBlock BuildTable(WTable t)
        {
            _tables++;

            var widths = new List<double>();
            var grid = t.GetFirstChild<TableGrid>();
            if (grid != null)
                foreach (var gc in grid.Elements<GridColumn>())
                    widths.Add((gc.Width?.Value is { } w && double.TryParse(w, NumberStyles.Any, CultureInfo.InvariantCulture, out var dw) ? dw : 0) * TwipsToPx);

            var rows = new List<MTableRow>();
            int rowIndex = 0;
            foreach (var tr in t.Elements<WTableRow>())
            {
                var cells = new List<MTableCell>();
                foreach (var tc in tr.Elements<WTableCell>())
                {
                    var tcPr = tc.TableCellProperties;
                    int span = tcPr?.GridSpan?.Val?.Value ?? 1;

                    string? vMerge = null;
                    var vm = tcPr?.VerticalMerge;
                    if (vm != null)
                        vMerge = vm.Val?.Value == MergedCellValues.Restart ? "restart" : "continue";

                    double? cellW = null;
                    if (tcPr?.TableCellWidth?.Width?.Value is { } cw
                        && double.TryParse(cw, NumberStyles.Any, CultureInfo.InvariantCulture, out var cwv))
                        cellW = cwv * TwipsToPx;

                    string? shading = null;
                    var fill = tcPr?.Shading?.Fill?.Value;
                    if (fill != null && fill != "auto") shading = "#" + fill;

                    var blocks = new List<Block>();
                    foreach (var el in tc.ChildElements)
                    {
                        if (el is Paragraph cp) blocks.Add(BuildParagraph(cp));
                        else if (el is WTable nested) blocks.Add(BuildTable(nested));
                    }

                    cells.Add(new MTableCell(blocks, span, vMerge, cellW, shading));
                }
                rows.Add(new MTableRow(cells, IsHeader: rowIndex == 0));
                rowIndex++;
            }

            return new TableBlock(widths, rows);
        }

        private static int? HeadingLevel(string? styleId, int? outlineLevel)
        {
            if (styleId != null)
            {
                var m = Regex.Match(styleId, @"^Heading(\d)$", RegexOptions.IgnoreCase);
                if (m.Success) return int.Parse(m.Groups[1].Value);
            }
            return outlineLevel is { } o ? o + 1 : null;
        }

        private static int CountWords(List<MRun> runs)
        {
            int words = 0;
            foreach (var r in runs)
                if (!string.IsNullOrEmpty(r.Text))
                    words += r.Text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Length;
            return words;
        }
    }

    // ---- Effective-formatting accumulators ------------------------------------------------

    /// <summary>Accumulates effective run formatting; each Apply only overrides present values.</summary>
    private sealed class RunFmt
    {
        private bool? _bold, _italic, _underline, _strike;
        private string? _font, _color, _highlight;
        private double? _sizeHalfPts;

        public RunFmt Clone() => new()
        {
            _bold = _bold, _italic = _italic, _underline = _underline, _strike = _strike,
            _font = _font, _color = _color, _highlight = _highlight, _sizeHalfPts = _sizeHalfPts,
        };

        public void Apply(OpenXmlElement? props)
        {
            if (props == null) return;

            if (props.GetFirstChild<Bold>() is { } b) _bold = b.Val?.Value ?? true;
            if (props.GetFirstChild<Italic>() is { } i) _italic = i.Val?.Value ?? true;
            if (props.GetFirstChild<Strike>() is { } st) _strike = st.Val?.Value ?? true;
            if (props.GetFirstChild<Underline>() is { } u)
                _underline = (u.Val?.Value ?? UnderlineValues.Single) != UnderlineValues.None;
            if (props.GetFirstChild<RunFonts>()?.Ascii?.Value is { } f) _font = f;
            if (props.GetFirstChild<FontSize>()?.Val?.Value is { } sz
                && double.TryParse(sz, NumberStyles.Any, CultureInfo.InvariantCulture, out var hp))
                _sizeHalfPts = hp;
            if (props.GetFirstChild<Color>()?.Val?.Value is { } c && c != "auto")
                _color = "#" + c;
            if (props.GetFirstChild<Highlight>()?.Val is { } hl)
                _highlight = hl.InnerText; // enum value name, e.g. "yellow"
        }

        public MRun ToRun(string text, string? href) => new(
            Text: text,
            Bold: _bold, Italic: _italic, Underline: _underline, Strike: _strike,
            Font: _font,
            SizePt: _sizeHalfPts is { } hp ? hp / 2.0 : null,
            Color: _color, Highlight: _highlight, Href: href);
    }

    /// <summary>Accumulates effective paragraph formatting.</summary>
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
                Align = jc == JustificationValues.Both ? "justify" : jc.ToString().ToLowerInvariant();

            var ind = pPr.GetFirstChild<Indentation>();
            if (ind?.Left?.Value is { } left
                && double.TryParse(left, NumberStyles.Any, CultureInfo.InvariantCulture, out var lv))
                IndentLeftTwips = lv;

            var sp = pPr.GetFirstChild<SpacingBetweenLines>();
            if (sp?.Before?.Value is { } bf
                && double.TryParse(bf, NumberStyles.Any, CultureInfo.InvariantCulture, out var bfv))
                SpacingBeforeTwips = bfv;
            if (sp?.After?.Value is { } af
                && double.TryParse(af, NumberStyles.Any, CultureInfo.InvariantCulture, out var afv))
                SpacingAfterTwips = afv;

            if (pPr.GetFirstChild<OutlineLevel>()?.Val?.Value is { } ol)
                OutlineLevel = ol;
        }
    }

    // ---- Style + numbering indices --------------------------------------------------------

    /// <summary>Indexes styles by id and resolves basedOn chains (root base first).</summary>
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
            chain.Reverse(); // apply base first, most-derived last
            return chain;
        }
    }

    private sealed record LevelInfo(bool Ordered, int Start, double? IndentLeftTwips);

    /// <summary>Resolves numId → abstractNumId → level (numFmt, start, indent).</summary>
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

            double? indLeft = null;
            var ind = lvl.PreviousParagraphProperties?.Indentation?.Left?.Value;
            if (ind != null && double.TryParse(ind, NumberStyles.Any, CultureInfo.InvariantCulture, out var lv))
                indLeft = lv;

            return new LevelInfo(ordered, start, indLeft);
        }
    }
}
