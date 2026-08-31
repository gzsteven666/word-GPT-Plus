import assert from 'node:assert/strict'

import { readSelectedInlineImage, WordImageReadError } from '../src/api/wordImageRuntime.ts'

const originalWord = globalThis.Word

const installWordMock = (context: unknown) => {
  globalThis.Word = {
    run: async (callback: (value: unknown) => unknown) => callback(context),
  } as unknown as typeof Word
}

try {
  let syncCalls = 0
  let writeCalls = 0
  const emptyPictures = {
    items: [] as unknown[],
    load: () => undefined,
  }
  const emptyContext = {
    document: {
      getSelection: () => ({ inlinePictures: emptyPictures }),
    },
    sync: async () => {
      syncCalls++
    },
  }
  installWordMock(emptyContext)
  await assert.rejects(
    () => readSelectedInlineImage(),
    (error: unknown) => error instanceof WordImageReadError && error.code === 'WORD_IMAGE_NOT_SELECTED',
  )
  assert.equal(syncCalls, 1)
  assert.equal(writeCalls, 0)

  const picture = {
    width: 320,
    height: 180,
    altTextTitle: 'test image',
    altTextDescription: 'V27 TEST',
    imageFormat: 'Png',
    load: () => undefined,
    getBase64ImageSrc: () => ({ value: 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAAB' }),
    delete: () => {
      writeCalls++
    },
  }
  const pictures = {
    items: [picture],
    load: () => undefined,
  }
  const context = {
    document: { getSelection: () => ({ inlinePictures: pictures }) },
    sync: async () => undefined,
  }
  installWordMock(context)
  const result = await readSelectedInlineImage()
  assert.deepEqual(result, {
    dataUrl: 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAAB',
    width: 320,
    height: 180,
    altTextTitle: 'test image',
    altTextDescription: 'V27 TEST',
    imageFormat: 'Png',
  })
  assert.equal(writeCalls, 0)

  const ambiguousPictures = {
    items: [picture, picture],
    load: () => undefined,
  }
  installWordMock({
    document: { getSelection: () => ({ inlinePictures: ambiguousPictures }) },
    sync: async () => undefined,
  })
  await assert.rejects(
    () => readSelectedInlineImage(),
    (error: unknown) => error instanceof WordImageReadError && error.code === 'WORD_IMAGE_AMBIGUOUS',
  )
  assert.equal(writeCalls, 0)
} finally {
  globalThis.Word = originalWord
}

console.log('word image runtime tests: PASS')
