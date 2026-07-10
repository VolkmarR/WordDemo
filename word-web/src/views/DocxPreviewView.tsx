import { useEffect, useRef, useState } from 'react'
import { renderAsync } from 'docx-preview'
import './DocxPreviewView.css'

// Approach B: fetch the raw .docx bytes and let docx-preview parse + render them entirely in
// the browser. It produces paginated, high-fidelity HTML (pages, columns, images, tables) and
// is render-only by design, so the result is inherently read-only. Nothing leaves the browser.
type Props = {
  file: File
}

export default function DocxPreviewView({ file }: Props) {
  const hostRef = useRef<HTMLDivElement>(null)
  const [status, setStatus] = useState('')

  useEffect(() => {
    let cancelled = false

    const host = hostRef.current
    if (!host)
      return

          ;(async () => {
      try {
        const formData = new FormData()
        formData.append('file', file)

        const res = await fetch('/api/document/file', {
          method: 'POST',
          body: formData,
        })

        if (!res.ok)
          throw new Error(`API responded with ${res.status}`)

        const blob = await res.blob()

        if (cancelled)
          return

        host.innerHTML = ''

        await renderAsync(blob, host, undefined, {
          className: 'docx',
          inWrapper: true,
          breakPages: true,
          experimental: true,
          useBase64URL: true,
        })

        if (!cancelled)
          setStatus('')
      }
      catch (err) {
        if (!cancelled)
          setStatus(
              err instanceof Error
                  ? err.message
                  : 'Failed to render document'
          )
      }
    })()

    return () => {
      cancelled = true
    }
  }, [file])

  return (
      <>
        {status && <div>{status}</div>}
        <div ref={hostRef} />
      </>
  )
}
