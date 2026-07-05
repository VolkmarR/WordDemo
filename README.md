# Word Preview POC

A **high-fidelity, read-only preview** of a Word document (`word-demo.docx`) in the browser,
rendered **four different ways behind a tab switcher** so the approaches can be compared on
fidelity, performance, and dependency footprint, plus a **fifth tab that converts the document
to PDF on the server** (two in-process engines). Scroll + text selection is the only
interaction; there are no editing affordances. A heading-outline sidebar and a word/paragraph
status line (the Word-flavored equivalents of the Excel POC's formula bar and status bar)
round out the backend tab.

**The Word file is never sent to any external service.** The backend serves it same-origin;
every approach parses it either on the .NET server or in the browser.

This mirrors the methodology of the sibling [Excel Preview POC](../ExcelDemo): one **shared
document model** produced on the backend, multiple renderers fed from it or from the raw bytes,
and OSS-only, permissively-licensed libraries (MIT / Apache-2.0 / BSD).

- `WordApi/` — ASP.NET Core (.NET 10) minimal API. Parses the `.docx` with the Open XML SDK
  into a flat `DocumentModel` and serves it as JSON, plus the raw bytes.
- `word-web/` — React 19 + Vite 8 (Rolldown) + React Compiler, pnpm. The tab switcher and all
  four renderers.
- `word-demo.docx` — the well-known WebAIM/IITAA accessibility "Sample Document": a Heading
  1/2/3 hierarchy (9 heading paragraphs), an ordered list with a nested bulleted sub-list, two
  inline images with alt text (`image1.gif`, `image2.png`), a simple table, a complex table
  with merged header cells, and a two-column section. It has **no** headers/footers, footnotes,
  or charts.

---

## The four approaches

| Tab | Parses | Renders with |
|-----|--------|--------------|
| **A — Backend · Open XML SDK** | .NET server (`DocumentFormat.OpenXml`) → shared `DocumentModel` JSON | hand-rolled React renderer (`src/doc/DocView.tsx`) |
| **B — Frontend · docx-preview** | browser (`docx-preview`) | docx-preview's own paginated HTML |
| **C — Frontend · mammoth** | browser (`mammoth`) | semantic HTML + a small readable stylesheet |
| **D — Backend model · Tiptap** | reuses **A's** JSON | off-the-shelf OSS editor (`@tiptap/react`, read-only) |

Approach A is the Word analog of the Excel POC's ClosedXML tab — except there is **no
higher-level "ClosedXML for Word"** with a compatible license, so it uses the low-level Open
XML SDK directly and resolves style inheritance, numbering, and images by hand. Approaches B
and C are browser-side parsers (analogous to the ExcelJS tab). Approach D is the "feed the same
backend model to an off-the-shelf component" tab (analogous to the react-data-grid / Jspreadsheet
tabs) — its point is to compare a hand-rolled renderer against an OSS one consuming identical data.

---

## Dependencies added per option

| Option | Library added | Direct deps | Transitive packages | Installed size | Shipped to the browser |
|--------|---------------|:-----------:|:-------------------:|---------------:|-----------------------:|
| **A — Backend · Open XML SDK** | `DocumentFormat.OpenXml` (NuGet, MIT) | **1** | 2 | **~8.3 MB** of DLLs (server-side) | **0 B** — the browser only receives compact JSON |
| **B — Frontend · docx-preview** | `docx-preview` (npm, Apache-2.0) | **1** (`jszip`) | 5 | **~1.9 MB** | **~171 KB** (49 KB gzip), lazy chunk |
| **C — Frontend · mammoth** | `mammoth` (npm, BSD-2-Clause) | **10** | 15 | **~6.9 MB** | **~492 KB** (119 KB gzip), lazy chunk |
| **D — Backend model · Tiptap** | `@tiptap/react` + `starter-kit` + `extension-image` + `extension-table` (npm, MIT) | **4** | ~49 | **~11 MB** | **~476 KB** (148 KB gzip), lazy chunk |

### Notes on the numbers

Measured with `pnpm list` / package-store directory sizes (`node_modules/.pnpm`) / built chunk
sizes from `pnpm build`, and `dotnet` `bin` output for the server DLLs.

- **A (Open XML SDK)** — `DocumentFormat.OpenXml` 3.5.1 pulls exactly two transitive packages
  (`DocumentFormat.OpenXml.Framework`, `System.IO.Packaging`). All parsing happens server-side;
  the browser receives only the resolved `DocumentModel` JSON — **~54 KB raw / ~16 KB gzip** for
  this document, most of it the two base64-embedded images. The custom renderer itself is part
  of the shared app bundle (`~201 KB / 64 KB gzip`, also used by D), not a per-approach library.
- **B (docx-preview)** — one dependency plus `jszip` (and jszip's tiny `lie`/`pako`/`setimmediate`
  chain). The smallest install of the three browser options and, byte-for-byte, the best fidelity.
- **C (mammoth)** — ten direct dependencies (`@xmldom/xmldom`, `bluebird`, `jszip`, `underscore`,
  `xmlbuilder`, …). Larger install and a larger browser chunk than docx-preview, yet it renders
  *less* (semantic HTML only) — by design.
- **D (Tiptap)** — the heaviest by far: four packages that fan out to ~49 `@tiptap/*` +
  `prosemirror-*` packages (~11 MB installed, ~476 KB shipped). It adds **no** parsing library —
  it consumes approach A's JSON — so its cost is purely the editor framework.

### Rough comparison

```
Browser payload:   A ~0 B*    B 171 KB    C 492 KB    D 476 KB      (*JSON only, ~16 KB gzip)
Packages added:    A 1 (+2)   B 1 (+5)    C 10 (+15)  D 4 (+~49)
Install footprint: A 8.3 MB   B 1.9 MB    C 6.9 MB    D 11 MB
```

**Takeaway:** for a *read-only, high-fidelity* preview, **docx-preview (B) wins on
fidelity-per-byte** — one small dependency, pages/columns/tables/images out of the box, nothing
shipped to any server. The **backend model (A)** is the choice when the browser payload must
stay near-zero or when the parsed structure is needed server-side anyway; it costs a hand-rolled
renderer but ships **0 B** of parsing library. **mammoth (C)** is the right tool only when you
*want* to discard presentation (content extraction, reflowable/semantic output). **Tiptap (D)**
is hard to justify for read-only preview — it is an editing framework, and you pay its full
weight for a document you never edit.

---

## Running it

**Prerequisites:** .NET SDK 10, Node 20+, and pnpm (`corepack enable`). On the **backend's host**,
the fonts used by the document (Calibri, or the metric-compatible `Carlito`) must be installed for
the QuestPDF PDF export to render real glyphs — see [Deploying](#deploying-linux-fonts--licensing).

```bash
# 1. Backend (serves the parsed model + the raw .docx on http://localhost:5269)
cd WordApi
dotnet run --launch-profile http
#   Swagger UI:  http://localhost:5269/swagger

# 2. Frontend (Vite dev server on http://localhost:5173, proxies /api → 5269)
cd word-web
pnpm install
pnpm dev
```

Open http://localhost:5173 and switch tabs. The Vite proxy forwards `/api` to the backend, so
no CORS configuration is needed. Point the backend at a different file with the `WordFile`
config key (default `../word-demo.docx`).

The API:

| Endpoint | Used by | Returns |
|----------|---------|---------|
| `GET /api/document` | A, D | the resolved `DocumentModel` as JSON |
| `GET /api/document/file` | B, C | the raw `.docx` bytes (same-origin) |
| `GET /api/document/pdf?engine=oss\|devexpress` | E | the document converted to PDF, in-process (`application/pdf`) |

---

## Fidelity checklist (results per tab)

Checked against `word-demo.docx`. ✅ correct · ⚠️ present but simplified · ❌ not rendered.

| Feature | A · Open XML SDK | B · docx-preview | C · mammoth | D · Tiptap |
|---------|:---:|:---:|:---:|:---:|
| H1 red `C00000` / H2 / H3 sizes & colors | ✅ | ✅ | ⚠️ sizes only, no color | ⚠️ sizes only, no color |
| Ordered list (6 items) 1–6 | ✅ | ✅ | ⚠️ restarts at "1. Columns" | ✅ (continues to 6) |
| Nested bulleted sub-list under item 5 | ✅ indented | ✅ | ⚠️ flat, not nested | ⚠️ separate list |
| Both images inline, correct size, alt text | ✅ | ✅ | ✅ | ✅ |
| Simple table: header row + borders | ✅ | ✅ | ✅ | ✅ |
| Complex table: merged header cells (`gridSpan`) | ✅ real colspans | ✅ | ✅ | ✅ |
| Two-column section flows as two columns | ✅ CSS multicol | ✅ | ❌ single column | ❌ single column |
| Run formatting (bold / underline / link color) | ✅ | ✅ | ⚠️ bold/links, no colors | ⚠️ bold/links, no colors |
| Whole document scrolls; no editing possible | ✅ | ✅ | ✅ | ✅ |

**Notes on the ⚠️/❌ cells:**

- **mammoth** deliberately emits *semantic* HTML: no fonts/colors/columns/page look. It also
  breaks the ordered list into two `<ol>`s around the interrupting bulleted list, so "Columns"
  restarts at 1 — an inherent limitation of a semantic converter, not a bug.
- **Tiptap** consumes approach A's model faithfully, but Tiptap's StarterKit has **no
  color/highlight/font marks**, so run colors (the red H1) and fonts are dropped; it has no
  multi-column layout. The ordered-list numbering *is* preserved (the converter sets the second
  list's `start` to 6). These are exactly the "OSS component fed the same model" trade-offs the
  tab exists to reveal.
- **A** and **B** both render the two-column section, merged table headers, the red H1, and the
  correctly-nested bulleted sub-list.

---

## Headings, lists & run formatting

| Option | How it's read | How it's drawn |
|--------|---------------|----------------|
| **A** | Backend resolves paragraph style → linked char style → direct formatting into flat effective runs; heading level from the style id / `outlineLvl`; list markers/indent resolved from `numbering.xml` (`numId → abstractNumId → level`) with ordered counters computed per `numId` while walking the body | `<h1>`–`<h3>` + `<p>` with inline styles from each run; list items as a hanging marker + body positioned by the resolved indent |
| **B** | docx-preview parses the OOXML in the browser | its own faithful HTML/CSS, closest to Word |
| **C** | mammoth maps OOXML → semantic elements | `<h1>`/`<ol>`/`<ul>`/`<strong>`; presentation intentionally dropped |
| **D** | converts A's JSON → ProseMirror doc (headings, `orderedList`/`bulletList` with `start`, marks) | Tiptap renders the ProseMirror doc read-only |

Numbering is the hardest part of approach A. The `%1.`-style format strings are simplified to
plain decimal/bullet markers (documented simplification); ordered counters are kept per `numId`
so the six numbered items stay 1–6 even though the two bulleted items interleave, and the
bulleted sub-list picks up the numbering level's deeper indent (1080 vs 720 twips) so it reads
as nested.

## Tables & merged cells

| Option | How merges are read | How they're drawn |
|--------|---------------------|-------------------|
| **A** | `gridSpan` (horizontal) and `vMerge` restart/continue (vertical) read per cell; grid column widths from `tblGrid` | real HTML `colSpan`/`rowSpan` (rowspans computed from `vMerge` runs); the first row is treated as the header |
| **B** | docx-preview | native table with correct spans |
| **C** | mammoth | `<table>` with `colspan`; borders from the readable stylesheet |
| **D** | A's JSON → Tiptap `table`/`tableRow`/`tableCell` nodes with `colspan`/`rowspan` | Tiptap table (borders are the component default) |

This document only uses **horizontal** merges (`gridSpan`) — the complex table's "May 2012" /
"September 2010" headers each span two columns. Vertical-merge (`vMerge`) handling is implemented
and exercised by the layout code but does not trigger on this file. Table borders are rendered
uniformly (the sample uses the `TableGrid` style = uniform 0.5 pt borders) rather than resolved
per edge — a documented simplification.

## Images

| Option | How images are read | How they're drawn |
|--------|---------------------|-------------------|
| **A** | backend pulls each `ImagePart`, base64-encodes it into a data URL, reads size from the drawing `extent` (EMU → px) and alt text from `docPr` | `<img>` with width/height + `alt` |
| **B** | docx-preview | `<img>` (data URLs) |
| **C** | mammoth inlines images as base64 data URLs | `<img>` with `alt` |
| **D** | A's data URLs → Tiptap `image` nodes | `<img>` |

Both images render inline at the right size with alt text present in the DOM in every tab. In
the source these are *floating/anchored* pictures; all four tabs place them inline in the text
flow (a simplification — no float/wrap).

## Columns, pages & what isn't in this document

| Feature | A | B | C | D | Notes |
|---------|:-:|:-:|:-:|:-:|-------|
| Two-column section | ✅ | ✅ | ❌ | ❌ | A uses CSS `column-count` per section + a forced `break-before: column` at the `<w:br type="column"/>` |
| True pagination / page look | ⚠️ | ✅ | ❌ | ⚠️ | A/D render one continuous sheet (page break shown as a divider); only docx-preview paginates |
| Headers / footers | — | — | — | — | none in this document |
| Footnotes | — | — | — | — | none in this document |
| Charts | — | — | — | — | none — the "chart" is a static PNG image (rendered as an image) |

---

## Known limitations

- **Approach A renders one continuous page, not real pages.** Section page geometry (size,
  margins) is taken from the first section; a page break is shown as a divider, not a new sheet.
  docx-preview (B) is the tab to use when true pagination matters.
- **Ordered-list markers are simplified** to decimal/bullet in A (the `numbering.xml` format
  strings and restart logic beyond per-`numId` counters are not fully emulated).
- **Table borders are uniform** in A/D (from the `TableGrid` style), not resolved per edge; cell
  shading is read but this document has none.
- **Tiptap (D) drops run colors, highlights, and fonts** (StarterKit has no such marks) and has
  no multi-column layout.
- **mammoth (C) drops all presentation** by design and splits the ordered list around the nested
  bulleted list.
- **Images are placed inline**, not floated/wrapped as in the source.
- This document has **no headers/footers, footnotes, or charts**, so those code paths are
  untested here.

---

## docx → PDF conversion (backend, in-process)

Tab **E** converts the document to PDF **on the .NET server, in-process** — no LibreOffice/Docker
sidecar, no separate service — and streams it back same-origin to the browser's native PDF viewer.
The `.docx` never leaves the machine, same contract as the other endpoints. Two engines are wired
behind a small `IPdfConverter` interface (`WordApi/Services/PdfConversion/`) and selected with
`?engine=`:

| Engine | `?engine=` | How | Fidelity | Status here |
|--------|-----------|-----|----------|-------------|
| **QuestPDF** (low-cost) | `oss` (default) | reuses approach A's parsed `DocumentModel` → QuestPDF layout; **paginates automatically** | good-enough business-doc | ✅ implemented, default |
| **DevExpress Office & PDF File API** (commercial) | `devexpress` | `RichEditDocumentServer.ExportToPdf` — direct `.docx` → PDF, one call | high | ⚙️ code present, behind the `DEVEXPRESS` build symbol (needs the DevExpress feed + license) |

### Requirements this was chosen against

- **In-process only** — pure managed .NET, no sidecar. **Linux containers** — so anything needing
  `libgdiplus`/GDI+ is out; SkiaSharp/HarfBuzz-based engines are preferred.
- **Good-enough** fidelity (content, headings, lists, tables, images correct; layout need not be
  pixel-identical to Word).
- **≤ $5,000/yr**, **11 developers**, **> $1M annual revenue** (so no free-tier eligibility).

### Candidate engines (2026 list prices)

Per-developer licenses count **developers who write/reference the library's API** — not the whole
company. A backend PDF feature is typically touched by a subset of the 11 devs, which is what keeps
the per-seat options viable.

| Engine | License model | Cost for this situation (>$1M rev, Linux) | Linux mechanism | Fidelity | In-process one-liner | Fits ≤$5k? |
|--------|---------------|-------------------------------------------|-----------------|:--------:|:--------------------:|:----------:|
| **QuestPDF** (our model → PDF) | Dual-licensed; **paid above $1M rev** — Professional $999 perpetual (≤10 devs) / Enterprise $2,999/yr (>10) | Bundled Skia (no libgdiplus) | good-enough | ⚠️ we build the layout | ✅ comfortably (paid) |
| **DevExpress Office File API** | Per-dev/yr | ~$1,199/dev new, ~$599 renew → 11 ≈ $13.2k; **$0 marginal if the team already owns DevExpress Universal** | `DevExpress.Drawing.Skia` | very good | ✅ `ExportToPdf` | ⚠️ only if ≤4 API-devs, or already owned |
| **GemBox.Document** | Perpetual, royalty-free deploy | Small Team **$4,450 (≤10 devs)** / Large Team $13,350 (≤50) | SkiaSharp + HarfBuzz | very good | ✅ `Save(".pdf")` | ✅ if ≤10 API-devs; ❌ at strict 11 |
| **Syncfusion DocIO** | Per-dev/yr (Community ineligible: >$1M & >5 devs) | ~$995/dev → 11 ≈ $10.9k | SkiaSharp | very good | ✅ `DocToPDFConverter` | ⚠️ only if ≤5 API-devs |
| **Aspose.Words** | Per-dev; or metered pay-per-use | ~$1,175 Small Business / ~$3,597 OEM per dev; or metered | Internal SkiaSharp | excellent | ✅ `Save(".pdf")` | ❌ upfront; ⚠️ metered for low volume |
| **Spire.Doc** | $999 dev / $2,999 OEM | needs **libgdiplus** on Linux → deployment friction | System.Drawing/GDI+ | good | ✅ `SaveToFile` | ❌ (Linux) |
| **Xceed Words** | $849.95 perpetual, fully managed | cheapest; lower fidelity on advanced features | Custom managed engine | decent | ✅ | ✅ budget / ⚠️ fidelity |

> Prices are 2026 **list** figures from public pages / marketplace summaries; verify with the
> vendor. Fidelity/Linux notes are corroborated by a public Word→PDF-on-Linux comparison of these
> libraries (see the DEV.to "Document Processing Libraries for Word to PDF on .NET 8 (Linux Azure
> Functions)" writeup).

### How the constraints narrow the field

- At **11 seats**, the per-developer engines (**DevExpress**, **Syncfusion**, **Aspose**) exceed
  $5k unless only a small subset of devs reference the API. **GemBox** sits right on the ≤10-dev
  boundary ($4,450 perpetual) — fine at ≤10 API-devs, over at a strict 11.
- **Spire.Doc** is ruled out by the Linux constraint (libgdiplus).
- **QuestPDF** is the only engine that fits ≤$5k *regardless* of how the seats are counted ($999
  perpetual for the backend team, or $2,999/yr Enterprise even if all 11 reference it) — but note
  it is **paid** here, not free, and you author the layout mapping yourself.

### Dependencies added per option

| Option | Library added | Direct deps | Transitive packages | Installed / deployed size |
|--------|---------------|:-----------:|:-------------------:|--------------------------:|
| **E · QuestPDF** | `QuestPDF` 2026.6.1 (NuGet, dual MIT/commercial) | **1** | **0** (bundles its own Skia) | package ~119 MB (native Skia for **all** RIDs); **deployed ~6 MB** per platform (`QuestPDF.dll` 0.5 MB + one `libQuestPdfSkia` ~5.5 MB) |
| **F · DevExpress** | `DevExpress.Document.Processor` + `DevExpress.Drawing.Skia` (private feed, commercial) | 2 | several DevExpress assemblies | *not installed here* (private feed); order of tens of MB |

### The two implemented approaches, head-to-head

Measured on `word-demo.docx` via `GET /api/document/pdf` (times exclude the process-start JIT; the
first request is ~1.1 s cold, warm requests ~30–40 ms):

| | **E · QuestPDF** (`oss`) | **F · DevExpress** (`devexpress`) |
|---|---|---|
| Output | 4 pages, ~74 KB | (enable the engine to measure) |
| Conversion time (warm) | ~35 ms | — |
| Pagination | ✅ automatic (real pages) | ✅ |
| Red H1 / heading sizes | ✅ | ✅ |
| Ordered list 1–6 + nested bullets | ✅ | ✅ |
| Simple + complex table (merged headers) | ✅ colspans render | ✅ |
| Inline images (pie chart, symbol) | ✅ | ✅ |
| Two-column section | ❌ single column (no native QuestPDF multicol) | ✅ |

Verified by rendering the endpoint output in Chrome's PDF viewer: the red **Sample Document** H1,
the 1–6 ordered list with the nested Simple/Complex-Tables bullets, the pie-chart image, and the
complex table whose "May 2012" / "September 2010" headers each span two columns all reproduce.
QuestPDF's one gap is the two-column section (rendered single-column).

### Recommendation & how to enable DevExpress

- **QuestPDF** is the recommended path here: cheapest full-featured option at 11 devs (though a
  **paid** license, not free), reuses the model we already parse, and matches the good-enough bar.
- For a **drop-in high-fidelity one-liner**, use **DevExpress** *if the team already owns a
  DevExpress subscription* (then it's $0 marginal), else **GemBox.Document** is the best-value
  perpetual option.
- The **DevExpress** converter (`DevExpressPdfConverter.cs`) compiles as a stub reporting HTTP 501
  until you: (1) add the DevExpress NuGet feed + your auth key via a `nuget.config`; (2) reference
  `DevExpress.Document.Processor` + `DevExpress.Drawing.Skia`; (3) register your license and define
  the `DEVEXPRESS` build symbol. GemBox drops into the same interface with no endpoint/frontend
  change (it's on nuget.org with an emailed trial key).

---

## Commercial alternatives (evaluation)

The four **preview** approaches are all OSS with permissive licenses. For completeness, here is
how the notable **commercial / copyleft** Word-in-the-browser options fare against the same
constraints as the Excel POC. (For server-side **PDF conversion** engines specifically, see the
dedicated section above.)

### Constraints

- Team of **11 developers**.
- Budget **≤ $5,000 / year**.
- **.NET + React SPA**, client-side interactive preview.
- **No server-side PDF/PNG rendering** of the document (the whole point is to render in the
  browser or from a compact model).
- Read-only preview (editing is a bonus, not a requirement).

### Vendors evaluated (2026 list prices)

| Product | Licensing model | Cost for 11 devs | Client-side, no server render | Renders .docx faithfully | Fits constraints |
|--------|-----------------|-----------------:|:-----------------------------:|:------------------------:|:----------------:|
| **Apryse WebViewer** (+ DOCX Editor add-on) | per-developer seat, custom quote | ~$2–5k **per seat** → well over budget for 11 | ✅ client-side | ✅ high | ❌ (price) |
| **Nutrient / PSPDFKit Web** (Document Editor) | component-based, custom quote | custom (historically $5k+ entry) | ✅ | ✅ | ❌ (price/opacity) |
| **Syncfusion DocumentEditor** (`ej2` / Blazor) | per-dev Team License; free Community tier | Community tier **excludes** this team (>5 devs / >10 employees); paid = custom quote | ✅ client-side | ✅ good | ❌ (Community ineligible; paid quote) |
| **TX Text Control DS Server** | **per-server** subscription, unlimited devs, yearly | one server license (historically $5k+/yr), renew ~40% | ⚠️ needs the DS Server backend | ✅ | ❌ (server component; price) |
| **SuperDoc** | dual: **AGPLv3** core + commercial Enterprise | commercial = contact `q@superdoc.dev` | ✅ client-side | ✅ excellent | ❌ (AGPL excluded; paid quote) |
| **OnlyOffice / Collabora Online** | **AGPLv3** core + commercial editions | commercial = custom | ⚠️ server-hosted document server | ✅ | ❌ (AGPL + server) |

> All prices are 2026 **list** figures gathered from public pages and marketplace summaries;
> most of these vendors quote per-customer, so treat the numbers as order-of-magnitude. Verify
> with the vendor before relying on them.

### How the constraints narrow the field

- **AGPL is out** (SuperDoc core, OnlyOffice/Collabora core) — same rule as the Excel POC.
- **Per-seat SDKs** (Apryse, Nutrient) blow the $5,000 budget almost immediately at 11 developers.
- **Syncfusion's** free Community License doesn't cover a team of 11 (limit is ≤ 5 developers /
  ≤ 10 employees), and the paid tier is a custom quote.
- **TX Text Control** and the commercial OnlyOffice/Collabora editions require a **server-side
  document backend**, which contradicts the "no server render / client-side" constraint.

### Conclusion

No commercial option clears the ≤ $5,000/year budget for 11 developers **and** stays purely
client-side without either an AGPL obligation or a server-side document service. The **$0
OSS fallback is the recommendation**: ship **docx-preview (B)** when out-of-the-box fidelity
(pages, columns, tables, images) matters most for the least weight, or the **backend Open XML
SDK model (A)** when the browser payload must stay near-zero or the parsed structure is useful
server-side. Both are permissively licensed and send nothing to any third party.

---

## Conventions & gotchas

- **Units.** EMU → px: `px = emu / 9525`. Twips (dxa) → px: `px = twips / 15` (1440 twips = 96 px
  = 1 in). Font size is half-points (`pt = sz / 2`).
- **Vite 8 / Rolldown.** Heavy tabs (docx-preview, mammoth, Tiptap) are `React.lazy`-loaded so
  each is an isolated chunk and the backend tab doesn't pay for them — the same chunk-isolation
  the Excel POC uses for Univer.
- **mammoth** is imported from its browser build (`mammoth/mammoth.browser`) to avoid Node
  built-ins; a tiny ambient `.d.ts` supplies types for that subpath.
- **The document sheets are intentionally always light** (a printed page is a light artifact),
  independent of the OS light/dark theme; the app chrome follows the system theme.
- **PDF status headers.** `/api/document/pdf` returns `X-Convert-Ms` and `X-Pdf-Bytes`; the tab-E
  status line reads them. They're exposed via `Access-Control-Expose-Headers` so they'd survive a
  cross-origin fetch too (in dev they're same-origin through the Vite proxy).
- **`NU1903` build warning** — `Microsoft.OpenApi` 2.0.0 (pulled transitively by the OpenAPI/Swagger
  packages) has a flagged advisory. It only affects the dev-time Swagger UI, not the PDF/parsing
  paths; bump `Microsoft.AspNetCore.OpenApi` when a patched version is available.

---

## Project layout

```
WordDemo/
├─ word-demo.docx                     the sample document
├─ WordApi/                           ASP.NET Core (.NET 10) minimal API
│  ├─ Program.cs                      endpoints + QuestPDF license + engine registry
│  ├─ Models/DocumentModel.cs         shared, style-resolved model (mirrors model.ts)
│  ├─ Services/DocumentReader.cs      Open XML SDK parser (styles, numbering, images)
│  └─ Services/PdfConversion/
│     ├─ IPdfConverter.cs             engine interface (Engine id + Convert(path))
│     ├─ QuestPdfConverter.cs         model → QuestPDF (engine "oss", default)
│     └─ DevExpressPdfConverter.cs    RichEditDocumentServer (engine "devexpress", behind DEVEXPRESS)
└─ word-web/                          React 19 + Vite 8, pnpm
   └─ src/
      ├─ App.tsx                       tab switcher (A–E)
      ├─ model.ts                      TS mirror of DocumentModel
      ├─ adapters/backendModel.ts      fetches /api/document (A, D)
      ├─ doc/DocView.tsx               approach A custom renderer
      └─ views/                        DocxPreviewView · MammothView · TiptapView · PdfView
```

---

## Deploying (Linux, fonts & licensing)

- **Fonts (QuestPDF).** QuestPDF rasterizes text with its bundled Skia and only has access to fonts
  present on the host. A stock Linux container has none of the document's fonts, so text renders as
  tofu (□) or falls back to a default. Install a metric-compatible set in the image, e.g. for
  Debian/Ubuntu: `apt-get install -y fonts-crosextra-carlito fonts-liberation` (Carlito ≈ Calibri,
  Liberation ≈ Arial/Times). Alternatively embed a `.ttf` and register it once at startup with
  `QuestPDF.Drawing.FontManager.RegisterFont(File.OpenRead("Calibri.ttf"))`. docx-preview/mammoth
  (browser tabs) are unaffected — they use the *viewer's* fonts.
- **QuestPDF license.** Set once at startup (`Program.cs`):
  `QuestPDF.Settings.License = LicenseType.Professional;` — use `Enterprise` if more than 10
  developers reference QuestPDF. This is a legal acknowledgement, not a key; forgetting it throws at
  first `GeneratePdf()` above the free-tier threshold.
- **No native GDI+ dependency.** Both wired engines are Linux-clean (QuestPDF's bundled Skia;
  DevExpress via `DevExpress.Drawing.Skia`), so **no `libgdiplus`** is required — the reason Spire.Doc
  was ruled out.

### Adding or swapping a PDF engine

The engine set is just an array in `Program.cs` keyed by `IPdfConverter.Engine`. To add one
(e.g. **GemBox.Document** as the commercial slot instead of DevExpress):

1. Add the NuGet package (GemBox is on nuget.org; DevExpress needs its private feed).
2. Implement `IPdfConverter` — `Engine => "gembox"`, and in `Convert(path)`:
   `ComponentInfo.SetLicense("FREE-LIMITED-KEY"); var doc = DocumentModel.Load(path); doc.Save(ms, new PdfSaveOptions()); return ms.ToArray();`
3. Add `new GemBoxPdfConverter()` to the `pdfConverters` array and the engine to `PdfView.tsx`'s
   `ENGINES` list. No endpoint change is needed.

To enable **DevExpress** specifically: add its NuGet feed + auth key via a `nuget.config`, reference
`DevExpress.Document.Processor` + `DevExpress.Drawing.Skia`, register your license, and add
`<DefineConstants>$(DefineConstants);DEVEXPRESS</DefineConstants>` to `WordApi.csproj` (or
`dotnet build -p:DefineConstants=DEVEXPRESS`).

---

## Troubleshooting

| Symptom | Cause / fix |
|---------|-------------|
| Tab A/D/E show *"Could not reach the document API"* | Backend isn't running or is on a different port. Start it with `dotnet run --launch-profile http` (listens on `:5269`, which the Vite proxy targets). |
| Tab E DevExpress shows *"engine is not enabled in this build"* (HTTP 501) | Expected until DevExpress is configured — see [Adding or swapping a PDF engine](#adding-or-swapping-a-pdf-engine). Use `?engine=oss` meanwhile. |
| PDF text is boxes/□ or the wrong font (esp. in a Linux container) | Missing fonts on the backend host — install Carlito/Liberation or register a `.ttf`; see [Deploying](#deploying-linux-fonts--licensing). |
| `GeneratePdf()` throws a QuestPDF license exception | `QuestPDF.Settings.License` not set (or set below your tier). Set it at startup. |
| `DocumentLayoutException: conflicting size constraints` from QuestPDF | Something exceeds the page width (e.g. constant column widths). The converter already uses proportional (`RelativeColumn`) table columns to avoid this; keep new content width-flexible. |
| First `/api/document/pdf` call is ~1 s, later ones ~30 ms | Expected — the first request pays QuestPDF/Skia JIT + init warm-up; steady-state is tens of ms. |
| `.docx` is locked/open in Word while the API reads it | `DocumentReader` opens with `FileShare.ReadWrite`, so this is fine; if you swap in an engine that opens exclusively, close Word first. |
