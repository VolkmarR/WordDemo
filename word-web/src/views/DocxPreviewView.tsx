import { useEffect, useRef, useState } from 'react'
import { renderAsync } from 'docx-preview'
import './DocxPreviewView.css'

// Approach B: fetch the raw .docx bytes and let docx-preview parse + render them entirely in
// the browser. It produces paginated, high-fidelity HTML (pages, columns, images, tables) and
// is render-only by design, so the result is inherently read-only. Nothing leaves the browser.

export default function DocxPreviewView() {
  const hostRef = useRef<HTMLDivElement>(null)
  const [status, setStatus] = useState<string>('Rendering…')

  useEffect(() => {
    let cancelled = false
    const host = hostRef.current
    if (!host) return
    ;(async () => {
      try {
        const res = await fetch('/api/document/file')
        if (!res.ok) throw new Error(`API responded with ${res.status}`)
        const blob = await res.blob()
        if (cancelled) return
        host.innerHTML = ''
        await renderAsync(blob, host, undefined, {
          className: 'docx',
          inWrapper: true,
          breakPages: true,
          experimental: true, // enables tab-stop handling
          useBase64URL: true,
        })
        if (!cancelled) setStatus('')
      } catch (err: unknown) {
        if (!cancelled)
          setStatus(err instanceof Error ? err.message : 'Failed to render document')
      }
    })()
    return () => {
      cancelled = true
    }
  }, [])

  return (
    <div className="dp-scroll">
      {status && <div className="msg">{status}</div>}
      <div className="dp-host" ref={hostRef} />
    </div>
  )
}
