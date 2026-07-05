import type { DocumentModel } from '../model'

/** Shared result shape carried by every adapter (mirrors the Excel POC's LoadResult). */
export interface LoadResult {
  model: DocumentModel
  ms: number
}

/** Approach A: fetch the server-parsed DocumentModel JSON from the .NET backend. */
export async function loadFromBackend(): Promise<LoadResult> {
  const t0 = performance.now()
  const res = await fetch('/api/document')
  if (!res.ok) throw new Error(`API responded with ${res.status}`)
  const model = (await res.json()) as DocumentModel
  return { model, ms: performance.now() - t0 }
}
