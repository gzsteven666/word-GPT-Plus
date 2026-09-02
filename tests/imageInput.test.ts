import assert from 'node:assert/strict'

import { HumanMessage } from '@langchain/core/messages'

import {
  appendImageAttachments,
  buildEphemeralMultimodalMessage,
  clearSentImageAttachments,
  constrainImageDimensions,
  getClipboardImageFiles,
  getImageCapabilityGate,
  type ImageAttachment,
  ImageInputError,
  isSupportedImageFile,
  MAX_IMAGE_ATTACHMENTS,
  MAX_IMAGE_BYTES,
  MAX_IMAGE_DIMENSION,
  MAX_IMAGE_TOTAL_PAYLOAD_BYTES,
  preventClipboardImageTextInsertion,
  resolveImageSendText,
  sanitizeHistoryMessage,
} from '../src/utils/imageInput.ts'

const file = (type: string, size = 1024) => new File([new Uint8Array(size)], 'sample.png', { type })

assert.equal(isSupportedImageFile(file('image/png')), true)
assert.equal(isSupportedImageFile(file('image/jpeg')), true)
assert.equal(isSupportedImageFile(file('image/webp')), true)
assert.equal(isSupportedImageFile(file('image/gif')), false)
assert.equal(isSupportedImageFile(file('image/png', MAX_IMAGE_BYTES + 1)), false)

const clipboardItem = (itemFile: File | null, kind = 'file', type = itemFile?.type || 'image/png') =>
  ({ kind, type, getAsFile: () => itemFile }) as unknown as DataTransferItem
const clipboardPng = file('image/png')
const clipboardJpeg = file('image/jpeg')
assert.deepEqual(
  getClipboardImageFiles({
    items: [clipboardItem(clipboardPng)],
    files: [],
  } as unknown as DataTransfer),
  [clipboardPng],
)
assert.deepEqual(
  getClipboardImageFiles({
    items: [clipboardItem(null, 'string', 'text/plain')],
    files: [clipboardJpeg],
  } as unknown as DataTransfer),
  [clipboardJpeg],
)
assert.deepEqual(
  getClipboardImageFiles({
    items: [clipboardItem(file('image/gif'))],
    files: [],
  } as unknown as DataTransfer),
  [],
)
assert.deepEqual(getClipboardImageFiles(null), [])
let clipboardDefaultPrevented = false
preventClipboardImageTextInsertion(
  {
    preventDefault: () => {
      clipboardDefaultPrevented = true
    },
  },
  [clipboardPng],
)
assert.equal(clipboardDefaultPrevented, true)

clipboardDefaultPrevented = false
preventClipboardImageTextInsertion(
  {
    preventDefault: () => {
      clipboardDefaultPrevented = true
    },
  },
  [],
)
assert.equal(clipboardDefaultPrevented, false)

assert.equal(resolveImageSendText('', true, 'analyze the attached image'), 'analyze the attached image')
assert.equal(resolveImageSendText('  describe this  ', true, 'analyze the attached image'), 'describe this')
assert.equal(resolveImageSendText('', false, 'analyze the attached image'), '')

assert.ok(MAX_IMAGE_DIMENSION > 0)
assert.deepEqual(constrainImageDimensions(8000, 4000), { width: MAX_IMAGE_DIMENSION, height: 800 })
assert.deepEqual(constrainImageDimensions(800, 400), { width: 800, height: 400 })

assert.equal(getImageCapabilityGate('yes'), 'allowed')
assert.equal(getImageCapabilityGate('no'), 'blocked')
assert.equal(getImageCapabilityGate('unknown'), 'unknown')

const attachment = (id: string, size: number, name = `${id}.png`): ImageAttachment => ({
  id,
  name,
  mimeType: 'image/png',
  size,
  width: 1,
  height: 1,
  dataUrl: `data:image/png;base64,${id}`,
})

const first = attachment('first', 100)
const second = attachment('second', 200)
const ephemeral = buildEphemeralMultimodalMessage('describe this', [first.dataUrl, second.dataUrl])
assert.ok(ephemeral instanceof HumanMessage)
assert.deepEqual(ephemeral.content, [
  { type: 'text', text: 'describe this' },
  { type: 'image_url', image_url: { url: first.dataUrl } },
  { type: 'image_url', image_url: { url: second.dataUrl } },
])

const historyMessage = sanitizeHistoryMessage(ephemeral)
assert.equal(historyMessage instanceof HumanMessage, true)
assert.equal(historyMessage.content, 'describe this')
assert.equal(JSON.stringify(historyMessage).includes('data:image'), false)
assert.equal(JSON.stringify(sanitizeHistoryMessage(new HumanMessage('plain'))).includes('data:image'), false)

const existing = [first]
const appended = appendImageAttachments(existing, [second])
assert.deepEqual(appended, [first, second])
assert.deepEqual(existing, [first])

assert.throws(
  () =>
    appendImageAttachments(
      Array.from({ length: MAX_IMAGE_ATTACHMENTS }, (_, index) => attachment(`item-${index}`, 1)),
      [second],
    ),
  (error: unknown) => error instanceof ImageInputError && error.code === 'IMAGE_COUNT_EXCEEDED',
)
assert.throws(
  () => appendImageAttachments([attachment('large', MAX_IMAGE_TOTAL_PAYLOAD_BYTES)], [second]),
  (error: unknown) => error instanceof ImageInputError && error.code === 'IMAGE_TOTAL_PAYLOAD_TOO_LARGE',
)

const sent = [first, second]
assert.deepEqual(clearSentImageAttachments(sent, sent, false), sent)
assert.deepEqual(clearSentImageAttachments(sent, [first], true), [second])

console.log('image input tests: PASS')
