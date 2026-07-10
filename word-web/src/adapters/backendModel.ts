import type {DocumentModel} from '../model'

/** Shared result shape carried by every adapter (mirrors the Excel POC's LoadResult). */
export interface LoadResult {
    model: DocumentModel
    ms: number
}

/** Approach A: fetch the server-parsed DocumentModel JSON from the .NET backend. */
export async function loadFromBackend(file: File): Promise<LoadResult> {
    const formData = new FormData()
    const API_URL = '/api/document'
    formData.append('file', file)

    const response = await fetch(API_URL, {
        method: 'POST',
        body: formData,
    })

    if (!response.ok)
        throw new Error(`HTTP ${response.status}`)

    return response.json()
}
