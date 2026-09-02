import assert from 'node:assert/strict'

import { MAX_PDF_FILE_BYTES, MAX_PDF_PAGES, PdfInputError, preparePdfAttachment } from '../src/utils/pdfInput.ts'

interface MockPage {
  getTextContent: () => Promise<{ items: { str?: string }[] }>
}

const makeFile = (name: string, size: number, type = 'application/pdf') =>
  new File([new Uint8Array(size)], name, { type })

const loadDocument = async () => ({
  numPages: 2,
  getPage: async (pageNumber: number): Promise<MockPage> => ({
    getTextContent: async () => ({ items: [{ str: pageNumber === 1 ? 'first page' : 'second page' }] }),
  }),
})

const attachment = await preparePdfAttachment(makeFile('notes.pdf', 12), loadDocument)
assert.equal(attachment.name, 'notes.pdf')
assert.equal(attachment.mimeType, 'application/pdf')
assert.equal(attachment.text, 'first page\nsecond page')
assert.equal(attachment.charCount, attachment.text.length)

await assert.rejects(
  () => preparePdfAttachment(new File(['hello'], 'notes.txt', { type: 'text/plain' }), loadDocument),
  (error: unknown) => error instanceof PdfInputError && error.code === 'PDF_TYPE_UNSUPPORTED',
)
await assert.rejects(
  () => preparePdfAttachment(makeFile('large.pdf', MAX_PDF_FILE_BYTES + 1), loadDocument),
  (error: unknown) => error instanceof PdfInputError && error.code === 'PDF_TOO_LARGE',
)
await assert.rejects(
  () =>
    preparePdfAttachment(makeFile('many-pages.pdf', 12), async () => ({
      numPages: MAX_PDF_PAGES + 1,
      getPage: async () => ({ getTextContent: async () => ({ items: [] }) }),
    })),
  (error: unknown) => error instanceof PdfInputError && error.code === 'PDF_PAGE_COUNT_EXCEEDED',
)
await assert.rejects(
  () =>
    preparePdfAttachment(makeFile('empty.pdf', 12), async () => ({
      numPages: 1,
      getPage: async () => ({ getTextContent: async () => ({ items: [{ str: '  ' }] }) }),
    })),
  (error: unknown) => error instanceof PdfInputError && error.code === 'PDF_EMPTY',
)

console.log('pdf input tests: PASS')
