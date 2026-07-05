import { useEffect, useState } from 'react'
import './PdfView.css'

// Approach E / F — the PDF is rendered on the .NET backend (in-process, same-origin) and shown
// here in the browser's native PDF viewer. We fetch it as a blob (rather than pointing the iframe
// straight at the URL) so we can read the X-Convert-Ms / X-Pdf-Bytes headers for the status line
// and reuse the same bytes for the Download button. Nothing leaves the machine.

type Engine = 'migradoc' | 'oss' | 'devexpress'

const ENGINES: { id: Engine; label: string; note: string }[] = [
  { id: 'migradoc', label: 'MigraDoc (free · MIT)', note: 'Independent Open XML walk → MigraDoc/PDFsharp. Own read+render pipeline, MIT-free.' },
  { id: 'oss', label: 'QuestPDF (low-cost)', note: "Backend DocumentModel → QuestPDF. Reuses approach A's parsed model." },
  { id: 'devexpress', label: 'DevExpress (commercial)', note: 'RichEditDocumentServer.ExportToPdf — direct .docx → PDF, higher fidelity.' },
]

interface Loaded {
  url: string
  ms: number | null
  bytes: number | null
}

export default function PdfView() {
  const [engine, setEngine] = useState<Engine>('oss')
  const [loaded, setLoaded] = useState<Loaded | null>(null)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    let cancelled = false
    let objectUrl: string | null = null
    setLoading(true)
    setError(null)

    fetch(`/api/document/pdf?engine=${engine}`)
      .then(async (res) => {
        if (!res.ok) {
          // 501 when the DevExpress engine isn't configured in this build; surface its message.
          // The body is an RFC 9110 ProblemDetails JSON — extract just the human-readable detail.
          const raw = await res.text().catch(() => '')
          let msg = raw
          try {
            const problem = JSON.parse(raw) as { detail?: string; title?: string }
            msg = problem.detail || problem.title || raw
          } catch {
            /* not JSON — use the raw text */
          }
          throw new Error(msg || `API responded with ${res.status}`)
        }
        const ms = Number(res.headers.get('X-Convert-Ms'))
        const bytes = Number(res.headers.get('X-Pdf-Bytes'))
        const blob = await res.blob()
        objectUrl = URL.createObjectURL(blob)
        if (cancelled) return
        setLoaded({
          url: objectUrl,
          ms: Number.isFinite(ms) ? ms : null,
          bytes: Number.isFinite(bytes) ? bytes : blob.size,
        })
      })
      .catch((err: unknown) => {
        if (!cancelled) setError(err instanceof Error ? err.message : 'Failed to render PDF')
      })
      .finally(() => {
        if (!cancelled) setLoading(false)
      })

    return () => {
      cancelled = true
      if (objectUrl) URL.revokeObjectURL(objectUrl)
    }
  }, [engine])

  const active = ENGINES.find((e) => e.id === engine)!

  return (
    <div className="pdf-layout">
      <div className="pdf-bar">
        <div className="pdf-engines" role="tablist" aria-label="PDF engine">
          {ENGINES.map((e) => (
            <button
              key={e.id}
              type="button"
              role="tab"
              aria-selected={e.id === engine}
              className={e.id === engine ? 'active' : ''}
              onClick={() => setEngine(e.id)}
            >
              {e.label}
            </button>
          ))}
        </div>
        <span className="pdf-note">{active.note}</span>
        <span className="pdf-status">
          {loading && 'Converting…'}
          {!loading && loaded && (
            <>
              {loaded.ms != null && `${loaded.ms} ms`}
              {loaded.bytes != null && ` · ${(loaded.bytes / 1024).toFixed(1)} KB`}
            </>
          )}
        </span>
        {loaded && !loading && (
          <a className="pdf-download" href={loaded.url} download={`word-demo-${engine}.pdf`}>
            Download PDF
          </a>
        )}
      </div>

      <div className="pdf-body">
        {error && (
          <div className="msg error">
            {engine === 'devexpress'
              ? 'DevExpress engine is not enabled in this build. '
              : 'Could not render the PDF. '}
            {error}
          </div>
        )}
        {!error && loaded && (
          <iframe className="pdf-frame" title={`PDF (${active.label})`} src={loaded.url} />
        )}
        {!error && !loaded && loading && <div className="msg">Converting document to PDF…</div>}
      </div>
    </div>
  )
}
