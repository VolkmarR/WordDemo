# Word Preview POC — Implementation Plan

Goal: render a **high-fidelity, read-only preview** of a Word file (`word-demo.docx`) in the
browser, following the same methodology as the Excel Preview POC in
`C:\Projects\WeBuild\ExcelDemo` — multiple approaches behind a **tab switcher**, one **shared
document model** where possible, and a README that compares the approaches on **fidelity,
performance, and dependency footprint**.

This plan is written to be executed by an agent working in `C:\Projects\WeBuild\WordDemo`.

---

## Requirements (same as the Excel POC)

1. **Read-only** preview — no editing affordances. Scroll + text selection is enough; no
   focused-cell concept here (that was spreadsheet-specific). A document **outline/heading
   navigation sidebar** is the Word-flavored equivalent of the formula bar — nice-to-have,
   see Phase 6.
2. Compare **backend parsing** vs **frontend parsing** approaches.
3. **OSS only, permissive licenses** (MIT / Apache-2.0 / BSD). No AGPL (excludes OnlyOffice,
   Collabora, SuperDoc core), no commercial tiers. Commercial options are documented as an
   *evaluation section* in the README at the end, not implemented (mirror the
   "Commercial alternatives" section of the Excel README).
4. The .docx **never leaves the machine**: the backend serves it same-origin; all parsing
   happens on the backend or in the browser.
5. Each approach's README entry must record: **library added, direct deps, transitive
   packages, installed size, bytes shipped to the browser** — same table format as
   `ExcelDemo/README.md`.

## Existing skeleton (do not scaffold from scratch)

- `WordApi/` — ASP.NET Core (.NET 10) minimal API, currently only the WeatherForecast
  template endpoint. Replace it.
- `word-web/` — React 19 + Vite 8 (rolldown) + React Compiler + oxlint, pnpm. `vite.config.ts`
  currently proxies `/weatherforecast` → `http://localhost:5269`; change the proxy to `/api`.
- `word-demo.docx` — the sample document (repo root). It is the well-known accessibility
  "Sample Document" containing: **Heading 1/2/3 hierarchy** (8 headings), an **ordered list
  with a nested bulleted list**, **two inline images with alt text** (`image1.gif`,
  `image2.png`), a **simple table** and a **complex table** (merged header cells), and a
  **two-column section**. It has **no** headers/footers, footnotes, or charts — note that in
  the README's limitations section. Bookmarks/internal links exist (`_top`).

## Approaches to implement (tabs, in this order)

### A — Backend · Open XML SDK → shared `DocumentModel` JSON + custom React renderer
- NuGet: `DocumentFormat.OpenXml` (MIT). This is the Word analog of ClosedXML in the Excel
  POC (there is no higher-level "ClosedXML for Word" with a compatible license — note this
  in the README).
- `WordApi/Models/DocumentModel.cs`: a JSON-serializable model of the document **body flow**:
  - `Block` discriminated union: `Paragraph` (runs, style id, heading level, numbering info
    with resolved list marker/level/ordered-ness, alignment, indent, spacing),
    `Table` (rows → cells → nested blocks, grid column widths, cell merges via
    `gridSpan`/`vMerge`, borders, shading), `Image` (base64 data URL, width/height in px from
    EMU, alt text) — images can also appear inline inside runs; model them as a run type.
  - `Run`: text, bold/italic/underline/strike, font, size (half-points → pt), color,
    highlight, hyperlink target.
  - Section info: page size/margins and **column count** per section (needed for the
    two-column part).
  - Resolve **style inheritance** on the backend (paragraph style → linked character style →
    direct formatting) so the frontend gets flat, effective formatting — same philosophy as
    the Excel POC where the backend ships resolved data and the browser gets compact JSON.
- `WordApi/Services/DocumentReader.cs`: walks `MainDocumentPart.Document.Body`, resolves
  styles from `StyleDefinitionsPart`, numbering from `NumberingDefinitionsPart`, images from
  `ImagePart`s.
- Endpoints in `Program.cs` (mirror `ExcelApi/Program.cs` including Swagger UI +
  OpenAPI annotations):
  - `GET /api/document` → `DocumentModel` JSON (path from config key `WordFile`, default
    `../word-demo.docx`).
  - `GET /api/document/file` → raw bytes,
    `application/vnd.openxmlformats-officedocument.wordprocessingml.document`.
- Frontend renderer `word-web/src/doc/DocView.tsx` (+ CSS): renders the model as a paged-look
  scrollable document (white sheet, page margins from section props; simple continuous flow —
  **no pagination**, document as one long page; note as limitation). Tables as HTML tables
  with `colspan`/`rowspan` from the model, lists with correct markers/indent, two-column
  section via CSS `columns`.

### B — Frontend · docx-preview (MIT)
- npm: `docx-preview` (MIT, ~no transitive deps besides `jszip`). Fetches
  `/api/document/file`, parses the .docx in the browser, and renders **paginated,
  high-fidelity HTML** into a container.
- This is the Word analog of the ExcelJS tab: 1 package, browser-side parsing, zero backend
  model. Expect the highest out-of-the-box fidelity (pages, columns, images, tables).
- Wrap in `word-web/src/views/DocxPreviewView.tsx`; render read-only (it is render-only by
  design), constrain with `breakPages: true`, `experimental: true` if needed for tab stops.

### C — Frontend · mammoth.js (BSD-2)
- npm: `mammoth` (browser build). Converts .docx → **semantic HTML** (headings, lists,
  tables, images as data URLs) — deliberately *low-fidelity* (no fonts/colors/columns/page
  look). Include it as the "how far does a semantic converter get" comparison point, styled
  with a small readable stylesheet.
- `word-web/src/views/MammothView.tsx`.

### D — Backend model · rendered by an OSS rich-text component (renderer swap)
- The Word analog of the react-data-grid / Jspreadsheet CE tabs: reuse **approach A's JSON**
  and swap only the renderer for an off-the-shelf OSS component, to compare "hand-rolled
  renderer" vs "OSS component fed the same model".
- Recommended: **Tiptap** (`@tiptap/react` + `@tiptap/starter-kit` + table/image extensions,
  MIT) in `editable: false` mode — convert `DocumentModel` → ProseMirror JSON. Alternative
  if Tiptap's transitive footprint looks bad in practice: `@lexical/react` (MIT, Meta).
  Record the dependency numbers either way; that comparison is the point of this tab.
- Known fidelity losses to document: no multi-column sections, approximate table borders,
  no page look. `word-web/src/views/TiptapView.tsx`.

If any approach turns out to be infeasible mid-implementation (e.g. library broken under
Vite 8/rolldown), keep the tab with a short "why not" note in the README rather than
silently dropping it — the Excel POC treats negative results as findings.

## Execution phases

1. **Backend** — replace the WeatherForecast template with the two `/api/document*`
   endpoints + `DocumentModel` + `DocumentReader` (approach A's parser). Add Swagger UI like
   `ExcelApi` (`Swashbuckle.AspNetCore.SwaggerUI` + `Microsoft.AspNetCore.OpenApi`).
   Verify with `dotnet run --launch-profile http` + `WordApi.http` requests.
2. **Frontend shell** — `App.tsx` tab switcher (same pattern as `ExcelDemo/excel-web/src/App.tsx`),
   `/api` proxy in `vite.config.ts`, shared TS types in `src/model.ts` mirroring
   `DocumentModel`, shared fetch adapter in `src/adapters/backendModel.ts`.
3. **Tab A** — custom `DocView` renderer against the checklist below.
4. **Tab B** — docx-preview. 5. **Tab C** — mammoth. 6. **Tab D** — Tiptap/Lexical fed by A's model.
5. **README.md** — write the comparison document in the same structure as
   `ExcelDemo/README.md`: intro, dependency table with measured sizes (use
   `pnpm list --depth Infinity` / package store sizes / built chunk sizes from
   `pnpm build`; if a chunk fails to build under rolldown, note it like the Excel README's
   Univer note), per-feature tables (images, tables, lists, columns), known limitations,
   run instructions, and a **commercial alternatives evaluation** (Apryse WebViewer,
   SuperDoc paid tier, Syncfusion DocumentEditor, TX Text Control, OnlyOffice/Collabora
   licensing) — research current list prices, apply the same constraints as the Excel POC
   (11 devs, ≤ $5,000/yr, client-side, no server PDF rendering) unless told otherwise.
6. **Optional polish** (only if all tabs work): heading-outline sidebar for Tab A (click to
   scroll), and a footer status line showing word/paragraph counts — the Word equivalents of
   the Excel POC's formula bar/status bar.

## Fidelity checklist (acceptance criteria per tab)

Every tab must be checked against `word-demo.docx` and results recorded in the README:

- [ ] Heading 1 ("Sample Document", red `C00000`) / Heading 2 / Heading 3 with correct sizes & colors
- [ ] Ordered list (6 items) with the nested bulleted sub-list under item 5, correct markers & indent
- [ ] Both images render inline with correct size; alt text present in the DOM
- [ ] Simple table: header row, borders
- [ ] Complex table: merged cells (`gridSpan`/`vMerge`) render as real spans
- [ ] Two-column section flows as two columns (or documented as a limitation)
- [ ] Run formatting: bold/italic/underline/color where present
- [ ] Whole document scrolls smoothly; no editing possible

## Conventions & gotchas

- **Run/verify:** backend `dotnet run --launch-profile http` from `WordApi/` (port 5269),
  frontend `pnpm install && pnpm dev` from `word-web/` (port 5173). Vite proxy makes CORS
  unnecessary. Keep both README run instructions identical in shape to the Excel one.
- Vite 8 uses the experimental **rolldown** bundler; if a library's UMD/CJS build misbehaves
  in dev, use `optimizeDeps.include`/`exclude` and document the workaround in
  `vite.config.ts` comments (see `ExcelDemo/excel-web/vite.config.ts` for the pattern; the
  ExcelDemo memory note: prod build can OOM on very large chunks — lazy-load heavy tabs with
  `React.lazy` so each approach is an isolated chunk, like the Excel POC does for Univer).
- EMU → px: `px = emu / 9525`. Word sizes: half-points for font size, twentieths of a point
  (dxa) for widths/indents (`px ≈ dxa / 20 * 96 / 72`).
- Numbering is the hardest part of approach A: resolve `numId → abstractNumId → level` and
  compute ordered-list counters **per numId instance** while walking the body; restart
  logic and `%1.` style format strings can be simplified to decimal/bullet for this POC —
  document any simplification.
- Commit style: this repo's history uses short imperative messages ("Added example with…");
  commit per phase.
