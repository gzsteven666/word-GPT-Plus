import type { TextFileAttachment } from './textFileInput'

export const MAX_PDF_FILE_BYTES = 10 * 1024 * 1024
export const MAX_PDF_PAGES = 100

export type PdfInputErrorCode =
  | 'PDF_TYPE_UNSUPPORTED'
  | 'PDF_TOO_LARGE'
  | 'PDF_PAGE_COUNT_EXCEEDED'
  | 'PDF_EMPTY'
  | 'PDF_PARSE_FAILED'

export class PdfInputError extends Error {
  readonly code: PdfInputErrorCode
  readonly fileName?: string

  constructor(code: PdfInputErrorCode, fileName?: string) {
    super(code)
    this.name = 'PdfInputError'
    this.code = code
    this.fileName = fileName
  }
}

interface PdfPage {
  getTextContent: () => Promise<{ items: { str?: string }[] }>
}

interface PdfDocument {
  numPages: number
  getPage: (pageNumber: number) => Promise<PdfPage>
}

export type PdfDocumentLoader = (data: Uint8Array) => Promise<PdfDocument>

const ensurePdfJsPromiseCompatibility = () => {
  const promiseConstructor = Promise as PromiseConstructor & {
    withResolvers?: <T>() => {
      promise: Promise<T>
      resolve: (value: T | PromiseLike<T>) => void
      reject: (reason?: unknown) => void
    }
  }
  if (typeof (promiseConstructor as { withResolvers?: unknown }).withResolvers === 'function') return
  promiseConstructor.withResolvers = <T>() => {
    let resolve!: (value: T | PromiseLike<T>) => void
    let reject!: (reason?: unknown) => void
    const promise = new Promise<T>((resolvePromise, rejectPromise) => {
      resolve = resolvePromise
      reject = rejectPromise
    })
    return { promise, resolve, reject }
  }
}

const defaultPdfDocumentLoader: PdfDocumentLoader = async data => {
  ensurePdfJsPromiseCompatibility()
  const { getDocument, GlobalWorkerOptions } = await import('pdfjs-dist/legacy/build/pdf.mjs')
  GlobalWorkerOptions.workerSrc = new URL('pdfjs-dist/legacy/build/pdf.worker.min.mjs', import.meta.url).toString()
  return getDocument({ data }).promise as Promise<PdfDocument>
}

const getExtension = (name: string): string => {
  const dot = name.lastIndexOf('.')
  return dot >= 0 ? name.slice(dot).toLowerCase() : ''
}

const createId = (): string => {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') return crypto.randomUUID()
  return `pdf-${Date.now()}-${Math.random().toString(36).slice(2)}`
}

export const preparePdfAttachment = async (
  file: File,
  loadDocument: PdfDocumentLoader = defaultPdfDocumentLoader,
): Promise<TextFileAttachment> => {
  if (getExtension(file.name) !== '.pdf' || (file.type !== '' && file.type !== 'application/pdf')) {
    throw new PdfInputError('PDF_TYPE_UNSUPPORTED', file.name)
  }
  if (file.size > MAX_PDF_FILE_BYTES) throw new PdfInputError('PDF_TOO_LARGE', file.name)

  let document: PdfDocument
  try {
    document = await loadDocument(new Uint8Array(await file.arrayBuffer()))
  } catch {
    throw new PdfInputError('PDF_PARSE_FAILED', file.name)
  }
  if (document.numPages > MAX_PDF_PAGES) throw new PdfInputError('PDF_PAGE_COUNT_EXCEEDED', file.name)

  try {
    const pages: string[] = []
    for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber += 1) {
      const content = await document.getPage(pageNumber).then(page => page.getTextContent())
      const pageText = content.items
        .map(item => item.str || '')
        .join('')
        .trim()
      if (pageText) pages.push(pageText)
    }
    const text = pages.join('\n').replace(/\r\n?/g, '\n').trim()
    if (!text) throw new PdfInputError('PDF_EMPTY', file.name)
    return {
      id: createId(),
      name: file.name,
      mimeType: 'application/pdf',
      size: file.size,
      charCount: text.length,
      text,
    }
  } catch (error) {
    if (error instanceof PdfInputError) throw error
    throw new PdfInputError('PDF_PARSE_FAILED', file.name)
  }
}
