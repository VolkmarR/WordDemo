import { useMemo } from 'react'
import { EditorContent, useEditor } from '@tiptap/react'
import { StarterKit } from '@tiptap/starter-kit'
import { Image } from '@tiptap/extension-image'
import { TableKit } from '@tiptap/extension-table'
import type { Block, DocumentModel, ParagraphBlock, Run, TableBlock } from '../model'
import './TiptapView.css'

// Approach D: reuse approach A's JSON model, but swap the hand-rolled renderer for an
// off-the-shelf OSS rich-text component (Tiptap, read-only). The model is converted to
// ProseMirror JSON here. Known fidelity losses (the point of the comparison): StarterKit has
// no color/highlight/font marks, so run colors (e.g. the red H1) and fonts are dropped; there
// is no multi-column layout and no page look; table borders are the component's defaults.

type PMNode = {
  type: string
  attrs?: Record<string, unknown>
  content?: PMNode[]
  marks?: { type: string; attrs?: Record<string, unknown> }[]
  text?: string
}

function marksOf(r: Run) {
  const m: { type: string; attrs?: Record<string, unknown> }[] = []
  if (r.bold) m.push({ type: 'bold' })
  if (r.italic) m.push({ type: 'italic' })
  if (r.underline) m.push({ type: 'underline' })
  if (r.strike) m.push({ type: 'strike' })
  if (r.href) m.push({ type: 'link', attrs: { href: r.href } })
  return m
}

function inlineContent(runs: Run[]): PMNode[] {
  const out: PMNode[] = []
  for (const r of runs) {
    if (r.break === 'line') {
      out.push({ type: 'hardBreak' })
      continue
    }
    if (r.break || r.image) continue // page/column breaks + images are not inline here
    const text = r.text ?? ''
    if (!text) continue
    const marks = marksOf(r)
    out.push(marks.length ? { type: 'text', text, marks } : { type: 'text', text })
  }
  return out
}

/** A paragraph → its Tiptap block(s): inline images become sibling block image nodes. */
function paragraphNodes(p: ParagraphBlock, asParagraph = false): PMNode[] {
  const nodes: PMNode[] = []
  for (const r of p.runs)
    if (r.image)
      nodes.push({ type: 'image', attrs: { src: r.image.src, alt: r.image.alt ?? null } })

  const inline = inlineContent(p.runs)
  if (p.headingLevel && !asParagraph) {
    nodes.push(
      inline.length
        ? { type: 'heading', attrs: { level: Math.min(p.headingLevel, 6) }, content: inline }
        : { type: 'heading', attrs: { level: Math.min(p.headingLevel, 6) } },
    )
  } else {
    nodes.push(inline.length ? { type: 'paragraph', content: inline } : { type: 'paragraph' })
  }
  return nodes
}

/** Render flags + rowspans from vertical merges (colspan comes straight from gridSpan). */
function tableLayout(t: TableBlock) {
  const rows = t.rows
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
      if (cell.vMerge === 'continue') info[ri][ci].render = false
      else if (cell.vMerge === 'restart') {
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

function tableNode(t: TableBlock): PMNode {
  const info = tableLayout(t)
  const rows: PMNode[] = t.rows.map((row, ri) => {
    const cells: PMNode[] = []
    row.cells.forEach((cell, ci) => {
      if (!info[ri][ci].render) return
      const content = cell.blocks
        .filter((b): b is ParagraphBlock => b.type === 'paragraph')
        .flatMap((b) => paragraphNodes(b, true))
      cells.push({
        type: row.isHeader ? 'tableHeader' : 'tableCell',
        attrs: { colspan: cell.gridSpan || 1, rowspan: info[ri][ci].rowSpan || 1 },
        content: content.length ? content : [{ type: 'paragraph' }],
      })
    })
    return { type: 'tableRow', content: cells }
  })
  return { type: 'table', content: rows }
}

function blocksToContent(blocks: Block[]): PMNode[] {
  const content: PMNode[] = []
  let i = 0
  while (i < blocks.length) {
    const b = blocks[i]
    if (b.type === 'paragraph' && b.list) {
      const ordered = b.list.ordered
      const start = parseInt(b.list.marker, 10)
      const items: PMNode[] = []
      while (i < blocks.length) {
        const lp = blocks[i]
        if (lp.type !== 'paragraph' || !lp.list || lp.list.ordered !== ordered) break
        items.push({ type: 'listItem', content: paragraphNodes(lp, true) })
        i++
      }
      content.push(
        ordered
          ? {
              type: 'orderedList',
              attrs: Number.isFinite(start) && start !== 1 ? { start } : {},
              content: items,
            }
          : { type: 'bulletList', content: items },
      )
      continue
    }
    if (b.type === 'table') {
      content.push(tableNode(b))
      i++
      continue
    }
    for (const n of paragraphNodes(b)) content.push(n)
    i++
  }
  return content
}

function toProseMirror(model: DocumentModel): PMNode {
  const blocks = model.sections.flatMap((s) => s.blocks)
  const content = blocksToContent(blocks)
  return { type: 'doc', content: content.length ? content : [{ type: 'paragraph' }] }
}

export default function TiptapView({ model }: { model: DocumentModel }) {
  const content = useMemo(() => toProseMirror(model), [model])
  const editor = useEditor({
    editable: false,
    immediatelyRender: false,
    extensions: [
      StarterKit,
      Image,
      TableKit.configure({ table: { resizable: false } }),
    ],
    content,
  })

  return (
    <div className="tt-scroll">
      <article className="tt-page">
        <EditorContent editor={editor} />
      </article>
    </div>
  )
}
