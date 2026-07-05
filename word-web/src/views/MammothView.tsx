import { useEffect, useRef, useState } from 'react'
import * as mammoth from 'mammoth/mammoth.browser'
import './MammothView.css'

// Approach C: mammoth converts the .docx to *semantic* HTML in the browser (headings, lists,
// tables, images as data URLs) and deliberately drops presentation — no fonts, colors,
// columns, or page look. It is the "how far does a plain semantic converter get?" comparison
// point, styled here with a small readable stylesheet.

export default function MammothView() {
  const hostRef = useRef<HTMLDivElement>(null)
  const [status, setStatus] = useState<string>('Converting…')

  useEffect(() => {
    let cancelled = false
    const host = hostRef.current
    if (!host) return
    ;(async () => {
      try {
        const res = await fetch('/api/document/file')
        if (!res.ok) throw new Error(`API responded with ${res.status}`)
        const arrayBuffer = await res.arrayBuffer()
        if (cancelled) return
        // mammoth inlines images as base64 data URLs by default.
        const result = await mammoth.convertToHtml({ arrayBuffer })
        if (cancelled) return
        host.innerHTML = result.value
        setStatus('')
      } catch (err: unknown) {
        if (!cancelled)
          setStatus(err instanceof Error ? err.message : 'Failed to convert document')
      }
    })()
    return () => {
      cancelled = true
    }
  }, [])

  return (
    <div className="mm-scroll">
      {status && <div className="msg">{status}</div>}
      <article className="mm-doc" ref={hostRef} />
    </div>
  )
}
