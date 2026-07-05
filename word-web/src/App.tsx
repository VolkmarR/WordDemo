import { lazy, Suspense, useEffect, useState } from 'react'
import './App.css'
import DocView from './doc/DocView'
import { loadFromBackend, type LoadResult } from './adapters/backendModel'

// docx-preview, mammoth and Tiptap are each large; load them as their own chunks so the
// backend tab doesn't pay for libraries it never uses (mirrors the Excel POC's React.lazy).
const DocxPreviewView = lazy(() => import('./views/DocxPreviewView'))
const MammothView = lazy(() => import('./views/MammothView'))
const TiptapView = lazy(() => import('./views/TiptapView'))
const PdfView = lazy(() => import('./views/PdfView'))

type Approach = 'a' | 'b' | 'c' | 'd' | 'e'

const TABS: { id: Approach; label: string; desc: string }[] = [
  {
    id: 'a',
    label: 'A — Backend · Open XML SDK',
    desc: 'Parsed on the .NET server (Open XML SDK) → shared DocumentModel JSON → custom React renderer.',
  },
  {
    id: 'b',
    label: 'B — Frontend · docx-preview',
    desc: 'Raw .docx parsed in the browser (docx-preview) → paginated, high-fidelity HTML. Nothing leaves the browser.',
  },
  {
    id: 'c',
    label: 'C — Frontend · mammoth',
    desc: 'Raw .docx converted to semantic HTML in the browser (mammoth.js) — deliberately low-fidelity.',
  },
  {
    id: 'd',
    label: 'D — Backend model · Tiptap',
    desc: "Approach A's JSON model mapped into an off-the-shelf OSS rich-text component (Tiptap, read-only).",
  },
  {
    id: 'e',
    label: 'E — Backend · PDF export',
    desc: 'Converted to PDF on the .NET server, in-process (MigraDoc, QuestPDF, or DevExpress), shown in the browser PDF viewer. Nothing leaves the machine.',
  },
]

function App() {
  const [approach, setApproach] = useState<Approach>('a')
  // The backend JSON model feeds tabs A and D; B and C parse the raw file themselves.
  const [backend, setBackend] = useState<LoadResult | null>(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)

  const needBackend = approach === 'a' || approach === 'd'

  useEffect(() => {
    if (!needBackend || backend) return
    let cancelled = false
    setLoading(true)
    setError(null)
    loadFromBackend()
      .then((res) => {
        if (!cancelled) setBackend(res)
      })
      .catch((err: unknown) => {
        if (!cancelled)
          setError(err instanceof Error ? err.message : 'Failed to load document')
      })
      .finally(() => {
        if (!cancelled) setLoading(false)
      })
    return () => {
      cancelled = true
    }
  }, [needBackend, backend])

  const tab = TABS.find((t) => t.id === approach)!

  function status() {
    if (needBackend && backend) {
      const s = backend.model.stats
      return `${Math.round(backend.ms)} ms · ${s.paragraphs} paragraphs · ${s.words} words · ${s.tables} tables · ${s.images} images`
    }
    return ''
  }

  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <strong>Word Preview POC</strong>
          <span className="desc">{tab.desc}</span>
        </div>
        <nav className="approach-tabs" role="tablist" aria-label="Rendering approach">
          {TABS.map((t) => (
            <button
              key={t.id}
              type="button"
              role="tab"
              aria-selected={t.id === approach}
              className={t.id === approach ? 'active' : ''}
              onClick={() => setApproach(t.id)}
            >
              {t.label}
            </button>
          ))}
        </nav>
        <div className="status">{status()}</div>
      </header>

      <main className="viewport">
        {needBackend && loading && <div className="msg">Loading document…</div>}
        {needBackend && error && (
          <div className="msg error">
            Could not reach the document API: {error}. Make sure the WordApi backend is
            running on http://localhost:5269.
          </div>
        )}

        {approach === 'a' && backend && <DocView model={backend.model} />}

        {approach === 'b' && (
          <Suspense fallback={<div className="msg">Loading docx-preview…</div>}>
            <DocxPreviewView />
          </Suspense>
        )}

        {approach === 'c' && (
          <Suspense fallback={<div className="msg">Loading mammoth…</div>}>
            <MammothView />
          </Suspense>
        )}

        {approach === 'd' && backend && (
          <Suspense fallback={<div className="msg">Loading Tiptap…</div>}>
            <TiptapView model={backend.model} />
          </Suspense>
        )}

        {approach === 'e' && (
          <Suspense fallback={<div className="msg">Loading PDF view…</div>}>
            <PdfView />
          </Suspense>
        )}
      </main>
    </div>
  )
}

export default App
