import assert from 'node:assert/strict'

import { AIMessage, HumanMessage } from '@langchain/core/messages'

import {
  DEFAULT_MAX_CONTEXT_CHARS,
  estimateMessageChars,
  exceedsContextBudget,
} from '../src/utils/conversationBudget.ts'
import { buildTextFileRequestText, type TextFileAttachment } from '../src/utils/textFileInput.ts'

assert.equal(estimateMessageChars([new HumanMessage('hello'), new AIMessage('world')]), 10)
assert.equal(exceedsContextBudget([new HumanMessage('hello')], 5), false)
assert.equal(exceedsContextBudget([new HumanMessage('hello')], 4), true)
assert.ok(DEFAULT_MAX_CONTEXT_CHARS > 0)

assert.equal(
  estimateMessageChars([
    {
      content: [
        { type: 'text', text: 'alpha' },
        { type: 'image_url', image_url: { url: 'ignored' } },
        { type: 'text', text: 'beta' },
      ],
    } as never,
  ]),
  9,
)

const textAttachment: TextFileAttachment = {
  id: 'text-1',
  name: 'notes.txt',
  mimeType: 'text/plain',
  size: 12,
  charCount: 12,
  text: 'attached text',
}
const textRequest = buildTextFileRequestText('summarize', [textAttachment])
assert.equal(estimateMessageChars([new HumanMessage(textRequest)]), textRequest.length)
assert.equal(estimateMessageChars([new HumanMessage('summarize')]), 'summarize'.length)

console.log('conversation budget tests: PASS')
