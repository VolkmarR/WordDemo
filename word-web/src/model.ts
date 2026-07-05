// Shared document model consumed by the tabs that use the backend (A and D).
// The backend (Open XML SDK) serializes to exactly this shape (camelCase); the Tiptap
// adapter maps out of it. Mirrors WordApi/Models/DocumentModel.cs field-for-field.

export type Align = 'left' | 'center' | 'right' | 'justify'

export interface InlineImage {
  src: string // base64 data URL
  widthPx: number
  heightPx: number
  alt: string | null
}

export interface Run {
  text?: string | null
  bold?: boolean | null
  italic?: boolean | null
  underline?: boolean | null
  strike?: boolean | null
  font?: string | null
  sizePt?: number | null
  color?: string | null // "#RRGGBB"
  highlight?: string | null
  href?: string | null // absolute URL or "#anchor"
  break?: 'line' | 'page' | 'column' | null
  image?: InlineImage | null
}

export interface ListInfo {
  ordered: boolean
  level: number
  marker: string // resolved display marker, e.g. "1." or "•"
}

export interface ParagraphBlock {
  type: 'paragraph'
  runs: Run[]
  styleId?: string | null
  headingLevel?: number | null // 1..3
  list?: ListInfo | null
  align?: Align | null
  indentLeftPx?: number | null
  spacingBeforePx?: number | null
  spacingAfterPx?: number | null
  anchor?: string | null // bookmark name, e.g. "_top"
}

export interface TableCell {
  blocks: Block[]
  gridSpan: number // colspan
  vMerge?: 'restart' | 'continue' | null
  widthPx?: number | null
  shading?: string | null // background fill hex
}

export interface TableRow {
  cells: TableCell[]
  isHeader: boolean
}

export interface TableBlock {
  type: 'table'
  colWidthsPx: number[]
  rows: TableRow[]
}

// Discriminated union on `type` so more block kinds can be added without reshaping.
export type Block = ParagraphBlock | TableBlock

export interface PageInfo {
  widthPx: number
  heightPx: number
  marginTopPx: number
  marginRightPx: number
  marginBottomPx: number
  marginLeftPx: number
}

export interface Section {
  columns: number
  page: PageInfo
  blocks: Block[]
}

export interface DocumentStats {
  paragraphs: number
  words: number
  headings: number
  tables: number
  images: number
}

export interface DocumentModel {
  sections: Section[]
  stats: DocumentStats
}

/** Plain text of a run list — used for word/heading extraction and outline labels. */
export function runsText(runs: Run[]): string {
  return runs.map((r) => r.text ?? '').join('')
}
