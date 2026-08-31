import assert from 'node:assert/strict'

import { AIMessage, HumanMessage } from '@langchain/core/messages'

import {
  DEFAULT_MAX_CONTEXT_CHARS,
  estimateMessageChars,
  exceedsContextBudget,
} from '../src/utils/conversationBudget.ts'

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

console.log('conversation budget tests: PASS')
