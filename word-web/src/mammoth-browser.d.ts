// mammoth ships types for its main entry but not for the browser subpath build. The browser
// build is a CJS module whose exports object carries the same functions; declare the subset used.
declare module 'mammoth/mammoth.browser' {
  interface Result {
    value: string
    messages: { type: string; message: string }[]
  }
  interface ArrayBufferInput {
    arrayBuffer: ArrayBuffer
  }
  export function convertToHtml(input: ArrayBufferInput, options?: unknown): Promise<Result>
  export function extractRawText(input: ArrayBufferInput): Promise<Result>
}
