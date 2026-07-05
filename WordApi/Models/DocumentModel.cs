using System.Text.Json.Serialization;

namespace WordApi.Models;

// Shape mirrors word-web/src/model.ts. Minimal APIs serialize with
// JsonSerializerDefaults.Web => camelCase on the wire.
//
// The backend (Open XML SDK) resolves style inheritance, numbering, and images so the
// frontend receives flat, effective formatting — same philosophy as the Excel POC where
// the backend ships resolved data and the browser gets compact JSON.

/// <summary>The whole document body flow, grouped into sections.</summary>
public record DocumentModel(
    List<Section> Sections,
    DocumentStats Stats);

/// <summary>Summary counts, used by the frontend status line.</summary>
public record DocumentStats(int Paragraphs, int Words, int Headings, int Tables, int Images);

/// <summary>
/// A section groups blocks and carries page geometry + column count. Word starts a new
/// section wherever a sectPr appears (in a paragraph's pPr, or as the body's final child).
/// The two-column part of the sample is its own section with <c>Columns == 2</c>.
/// </summary>
public record Section(
    int Columns,
    PageInfo Page,
    List<Block> Blocks);

/// <summary>Page size + margins in CSS pixels (twips / 15, i.e. 1440 twips = 96px = 1in).</summary>
public record PageInfo(
    double WidthPx, double HeightPx,
    double MarginTopPx, double MarginRightPx, double MarginBottomPx, double MarginLeftPx);

// ---- Block discriminated union: { "type": "paragraph" | "table", ... } -------------------
// Inline images live on runs (see Run.Image) because in .docx a picture is anchored inside a
// run; there is no block-level image in this sample. Documented in the README.

[JsonPolymorphic(TypeDiscriminatorPropertyName = "type")]
[JsonDerivedType(typeof(ParagraphBlock), "paragraph")]
[JsonDerivedType(typeof(TableBlock), "table")]
public abstract record Block;

/// <summary>A paragraph: a flow of runs with resolved paragraph-level formatting.</summary>
public record ParagraphBlock(
    List<Run> Runs,
    string? StyleId = null,
    int? HeadingLevel = null,       // 1..3 (from style / outlineLvl), null when not a heading
    ListInfo? List = null,          // set when the paragraph is a numbered/bulleted list item
    string? Align = null,           // "left" | "center" | "right" | "justify"
    double? IndentLeftPx = null,
    double? SpacingBeforePx = null,
    double? SpacingAfterPx = null,
    string? Anchor = null           // bookmark name (e.g. "_top") for internal-link targets
) : Block;

/// <summary>Resolved list marker for a paragraph (numId → abstractNum → level).</summary>
public record ListInfo(
    bool Ordered,
    int Level,          // ilvl (0-based)
    string Marker);     // resolved display marker, e.g. "1." or "•"

/// <summary>A table: grid column widths + rows of cells (cells hold nested blocks).</summary>
public record TableBlock(
    List<double> ColWidthsPx,
    List<TableRow> Rows) : Block;

public record TableRow(List<TableCell> Cells, bool IsHeader);

/// <summary>
/// A table cell. <see cref="GridSpan"/> is the colspan; <see cref="VMerge"/> is
/// "restart" | "continue" | null for vertical merges (the renderer computes rowspans from
/// the run of "continue" cells below a "restart").
/// </summary>
public record TableCell(
    List<Block> Blocks,
    int GridSpan = 1,
    string? VMerge = null,
    double? WidthPx = null,
    string? Shading = null);        // background fill hex, e.g. "#D9D9D9"

// ---- Runs --------------------------------------------------------------------------------

/// <summary>An inline run: text with effective character formatting, or a break, or an image.</summary>
public record Run(
    string? Text = null,
    bool? Bold = null, bool? Italic = null, bool? Underline = null, bool? Strike = null,
    string? Font = null,
    double? SizePt = null,          // points (half-points / 2)
    string? Color = null,           // "#RRGGBB"
    string? Highlight = null,       // highlight color name/hex
    string? Href = null,            // hyperlink target: absolute URL or "#anchor"
    string? Break = null,           // "line" | "page" | "column"
    InlineImage? Image = null);

/// <summary>An inline picture, embedded as a base64 data URL, sized in CSS pixels (EMU / 9525).</summary>
public record InlineImage(string Src, double WidthPx, double HeightPx, string? Alt);
