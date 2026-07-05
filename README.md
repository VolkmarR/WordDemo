# Word Preview POC

A **high-fidelity, read-only preview** of a Word document (`word-demo.docx`) in the browser,
rendered **four different ways behind a tab switcher** so the approaches can be compared on
fidelity, performance, and dependency footprint. Scroll + text selection is the only
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

## Commercial alternatives (evaluation)

The four implemented approaches are all OSS with permissive licenses. For completeness, here is
how the notable **commercial / copyleft** Word-in-the-browser options fare against the same
constraints as the Excel POC.

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
