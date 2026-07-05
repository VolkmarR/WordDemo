import { Fragment, useMemo, type CSSProperties, type JSX } from 'react'
import {
  runsText,
  type Block,
  type DocumentModel,
  type ParagraphBlock,
  type Run,
  type Section,
  type TableBlock,
} from '../model'
import './DocView.css'

/** Stable element id for a heading paragraph (its bookmark, else derived from its position). */
function headingId(p: ParagraphBlock, key: string): string | undefined {
  if (p.anchor) return p.anchor
  if (p.headingLevel) return `h-${key}`
  return undefined
}

// Custom renderer for approach A. Consumes the flat, style-resolved DocumentModel and draws a
// paged-look, read-only document: one white sheet (continuous flow — no pagination), page
// margins from the section properties, tables as real HTML tables with colspan/rowspan, lists
// with resolved markers/indent, and the two-column section via CSS multicol.

/** Effective character formatting → inline CSS. */
function runStyle(r: Run): CSSProperties {
  const s: CSSProperties = {}
  if (r.bold) s.fontWeight = 700
  if (r.italic) s.fontStyle = 'italic'
  const deco: string[] = []
  if (r.underline) deco.push('underline')
  if (r.strike) deco.push('line-through')
  if (deco.length) s.textDecoration = deco.join(' ')
  if (r.color) s.color = r.color
  if (r.font) s.fontFamily = `"${r.font}", var(--doc-font)`
  if (r.sizePt) s.fontSize = `${r.sizePt}pt`
  if (r.highlight && r.highlight !== 'none') s.backgroundColor = r.highlight
  return s
}

function renderRuns(runs: Run[], keyPrefix: string): JSX.Element[] {
  const out: JSX.Element[] = []
  runs.forEach((r, i) => {
    const key = `${keyPrefix}-${i}`
    if (r.image) {
      out.push(
        <img
          key={key}
          className="doc-img"
          src={r.image.src}
          alt={r.image.alt ?? ''}
          width={Math.round(r.image.widthPx) || undefined}
          height={Math.round(r.image.heightPx) || undefined}
        />,
      )
      return
    }
    // Column breaks are handled by paragraph splitting; a stray one falls through as nothing.
    if (r.break === 'column') return
    if (r.break === 'line' || r.break === 'page') {
      out.push(<br key={key} />)
      return
    }
    const text = r.text ?? ''
    const style = runStyle(r)
    if (r.href) {
      out.push(
        <a key={key} className="doc-link" href={r.href} style={style}>
          {text}
        </a>,
      )
      return
    }
    out.push(
      <span key={key} style={style}>
        {text}
      </span>,
    )
  })
  return out
}

function paragraphSpacing(p: ParagraphBlock): CSSProperties {
  const s: CSSProperties = {}
  if (p.align) s.textAlign = p.align
  if (p.spacingBeforePx != null) s.marginTop = p.spacingBeforePx
  if (p.spacingAfterPx != null) s.marginBottom = p.spacingAfterPx
  return s
}

function renderParagraph(p: ParagraphBlock, key: string): JSX.Element {
  // A paragraph that is only a page break: draw a visual page divider (no real pagination).
  if (p.runs.length > 0 && p.runs.every((r) => r.break === 'page')) {
    return <div key={key} className="doc-pagebreak" aria-hidden="true" />
  }

  // List item: hanging marker + body, positioned by the resolved indent.
  if (p.list) {
    const style: CSSProperties = { ...paragraphSpacing(p) }
    if (p.indentLeftPx != null) style.marginLeft = p.indentLeftPx
    return (
      <div key={key} className="doc-li" style={style}>
        <span className="doc-li-marker" aria-hidden="true">
          {p.list.marker}
        </span>
        <span className="doc-li-body">{renderRuns(p.runs, key)}</span>
      </div>
    )
  }

  const style: CSSProperties = { ...paragraphSpacing(p) }
  if (p.indentLeftPx != null) style.marginLeft = p.indentLeftPx

  const id = headingId(p, key)
  const Tag = (p.headingLevel ? `h${Math.min(p.headingLevel, 6)}` : 'p') as keyof JSX.IntrinsicElements
  const className = p.headingLevel ? 'doc-h' : 'doc-p'

  // Split on column breaks so the two-column section flows across CSS columns. Each fragment
  // after the first forces a new column with break-before.
  const segments: Run[][] = [[]]
  for (const r of p.runs) {
    if (r.break === 'column') segments.push([])
    else segments[segments.length - 1].push(r)
  }

  if (segments.length === 1) {
    return (
      <Tag key={key} id={id} className={className} style={style}>
        {renderRuns(p.runs, key)}
      </Tag>
    )
  }

  return (
    <Fragment key={key}>
      {segments.map((seg, i) => (
        <Tag
          key={`${key}-s${i}`}
          id={i === 0 ? id : undefined}
          className={className}
          style={i === 0 ? style : { ...style, breakBefore: 'column' }}
        >
          {renderRuns(seg, `${key}-s${i}`)}
        </Tag>
      ))}
    </Fragment>
  )
}

/** Precompute render flags + rowspans from vertical merges (gridSpan handles colspan directly). */
function tableLayout(t: TableBlock) {
  const rows = t.rows
  // Starting grid-column of each cell (accounting for gridSpan).
  const starts = rows.map((row) => {
    let g = 0
    return row.cells.map((c) => {
      const s = g
      g += c.gridSpan
      return s
    })
  })
  const info = rows.map((row) => row.cells.map(() => ({ render: true, rowSpan: 1 })))
  for (let ri = 0; ri < rows.length; ri++) {
    for (let ci = 0; ci < rows[ri].cells.length; ci++) {
      const cell = rows[ri].cells[ci]
      if (cell.vMerge === 'continue') {
        info[ri][ci].render = false
        continue
      }
      if (cell.vMerge === 'restart') {
        const g = starts[ri][ci]
        let span = 1
        for (let rj = ri + 1; rj < rows.length; rj++) {
          const idx = starts[rj].findIndex(
            (sc, k) => sc === g && rows[rj].cells[k].vMerge === 'continue',
          )
          if (idx >= 0) span++
          else break
        }
        info[ri][ci].rowSpan = span
      }
    }
  }
  return info
}

function renderTable(t: TableBlock, key: string): JSX.Element {
  const info = tableLayout(t)
  return (
    <table key={key} className="doc-table">
      {t.colWidthsPx.length > 0 && (
        <colgroup>
          {t.colWidthsPx.map((w, i) => (
            <col key={i} style={{ width: w }} />
          ))}
        </colgroup>
      )}
      <tbody>
        {t.rows.map((row, ri) => (
          <tr key={ri}>
            {row.cells.map((cell, ci) => {
              if (!info[ri][ci].render) return null
              const cellStyle: CSSProperties = {}
              if (cell.shading) cellStyle.background = cell.shading
              const inner = cell.blocks.map((b, bi) => renderBlock(b, `${key}-${ri}-${ci}-${bi}`))
              const common = {
                colSpan: cell.gridSpan > 1 ? cell.gridSpan : undefined,
                rowSpan: info[ri][ci].rowSpan > 1 ? info[ri][ci].rowSpan : undefined,
                style: cellStyle,
              }
              return row.isHeader ? (
                <th key={ci} scope="col" {...common}>
                  {inner}
                </th>
              ) : (
                <td key={ci} {...common}>
                  {inner}
                </td>
              )
            })}
          </tr>
        ))}
      </tbody>
    </table>
  )
}

function renderBlock(b: Block, key: string): JSX.Element {
  return b.type === 'paragraph' ? renderParagraph(b, key) : renderTable(b, key)
}

function renderSection(s: Section, i: number): JSX.Element {
  const style: CSSProperties = {}
  if (s.columns > 1) {
    style.columnCount = s.columns
    style.columnGap = 36
    style.columnRule = '1px solid var(--doc-rule)'
  }
  return (
    <section key={i} className="doc-section" style={style}>
      {s.blocks.map((b, bi) => renderBlock(b, `s${i}-${bi}`))}
    </section>
  )
}

interface OutlineItem {
  id: string
  level: number
  text: string
}

function scrollToHeading(id: string) {
  document.getElementById(id)?.scrollIntoView({ behavior: 'smooth', block: 'start' })
}

export default function DocView({ model }: { model: DocumentModel }) {
  // The sheet geometry comes from the first section (margins are uniform in this sample; a
  // single continuous page is used, so per-section page changes are collapsed — a limitation).
  const page = model.sections[0]?.page
  const pageStyle: CSSProperties = page
    ? {
        width: page.widthPx,
        paddingTop: page.marginTopPx,
        paddingRight: page.marginRightPx,
        paddingBottom: page.marginBottomPx,
        paddingLeft: page.marginLeftPx,
      }
    : {}

  // Heading-outline navigation — the Word-flavored equivalent of the Excel POC's formula bar.
  const outline = useMemo<OutlineItem[]>(() => {
    const items: OutlineItem[] = []
    model.sections.forEach((s, i) =>
      s.blocks.forEach((b, bi) => {
        if (b.type === 'paragraph' && b.headingLevel) {
          const text = runsText(b.runs).trim()
          if (text) items.push({ id: headingId(b, `s${i}-${bi}`)!, level: b.headingLevel, text })
        }
      }),
    )
    return items
  }, [model])

  const { stats } = model

  return (
    <div className="doc-layout">
      <nav className="doc-outline" aria-label="Document outline">
        <div className="doc-outline-title">Outline</div>
        {outline.map((it) => (
          <button
            key={it.id}
            type="button"
            className={`lvl${it.level}`}
            onClick={() => scrollToHeading(it.id)}
          >
            {it.text}
          </button>
        ))}
      </nav>

      <div className="doc-main">
        <div className="doc-scroll">
          <article className="doc-page" style={pageStyle}>
            {model.sections.map(renderSection)}
          </article>
        </div>
        <footer className="doc-footer">
          <span>{stats.paragraphs} paragraphs</span>
          <span>{stats.words} words</span>
          <span>{stats.headings} headings</span>
          <span>
            {stats.tables} tables · {stats.images} images
          </span>
        </footer>
      </div>
    </div>
  )
}
