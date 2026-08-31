import assert from 'node:assert/strict'

import { HumanMessage } from '@langchain/core/messages'

import {
  buildEphemeralMultimodalMessage,
  constrainImageDimensions,
  getImageCapabilityGate,
  isSupportedImageFile,
  MAX_IMAGE_BYTES,
  MAX_IMAGE_DIMENSION,
  sanitizeHistoryMessage,
} from '../src/utils/imageInput.ts'

const file = (type: string, size = 1024) => new File([new Uint8Array(size)], 'sample.png', { type })

assert.equal(isSupportedImageFile(file('image/png')), true)
assert.equal(isSupportedImageFile(file('image/jpeg')), true)
assert.equal(isSupportedImageFile(file('image/webp')), true)
assert.equal(isSupportedImageFile(file('image/gif')), false)
assert.equal(isSupportedImageFile(file('image/png', MAX_IMAGE_BYTES + 1)), false)
assert.ok(MAX_IMAGE_DIMENSION > 0)
assert.deepEqual(constrainImageDimensions(8000, 4000), { width: MAX_IMAGE_DIMENSION, height: 800 })
assert.deepEqual(constrainImageDimensions(800, 400), { width: 800, height: 400 })

assert.equal(getImageCapabilityGate('yes'), 'allowed')
assert.equal(getImageCapabilityGate('no'), 'blocked')
assert.equal(getImageCapabilityGate('unknown'), 'unknown')

const ephemeral = buildEphemeralMultimodalMessage('describe this', 'data:image/png;base64,abc')
assert.ok(ephemeral instanceof HumanMessage)
assert.deepEqual(ephemeral.content, [
  { type: 'text', text: 'describe this' },
  { type: 'image_url', image_url: { url: 'data:image/png;base64,abc' } },
])

const historyMessage = sanitizeHistoryMessage(ephemeral)
assert.equal(historyMessage instanceof HumanMessage, true)
assert.equal(historyMessage.content, 'describe this')
assert.equal(JSON.stringify(historyMessage).includes('data:image'), false)
assert.equal(JSON.stringify(sanitizeHistoryMessage(new HumanMessage('plain'))).includes('data:image'), false)

console.log('image input tests: PASS')
