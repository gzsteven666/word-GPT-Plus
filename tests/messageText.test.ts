import assert from 'node:assert/strict'

import { stripToolActivity } from '../src/utils/messageText.ts'

assert.equal(
  stripToolActivity('Answer\n\n🔧 Calling tool: read_document_nodes...\n✅ Tool read_document_nodes completed'),
  'Answer',
)
assert.equal(stripToolActivity('Answer\n\nMore detail'), 'Answer\n\nMore detail')

console.log('message text tests: PASS')
