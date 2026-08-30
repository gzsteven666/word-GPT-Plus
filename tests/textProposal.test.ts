import assert from 'node:assert/strict'

import { AppError } from '../src/api/errors.ts'
import { buildTextDiff, createTextChangeProposal, getProtectedObjects, hashText } from '../src/utils/textProposal.ts'

const before = '你好，世界。'
const after = '你好，新的世界。'
const diff = buildTextDiff(before, after)
assert.ok(diff.some(line => line.type === 'added' && line.text.includes('新的')))
assert.ok(diff.some(line => line.type === 'unchanged' && line.text.includes('你好')))
assert.equal(hashText(before), hashText(before))
assert.notEqual(hashText(before), hashText(after))

const proposal = createTextChangeProposal({
  operation: 'replace',
  beforeText: before,
  afterText: after,
  beforeOoxml: '<w:p><w:r><w:t>你好，世界。</w:t></w:r></w:p>',
  source: 'quick-action',
})
assert.equal(proposal.scope, 'selection')
assert.equal(proposal.status, 'pending')
assert.equal(proposal.expectedBeforeHash, hashText(before))

const objects = getProtectedObjects('<w:tbl><w:tr /></w:tbl><w:drawing />')
assert.deepEqual(objects, ['图片或绘图', '表格'])
assert.throws(
  () =>
    createTextChangeProposal({
      operation: 'replace',
      beforeText: 'table',
      afterText: 'text',
      beforeOoxml: '<w:tbl />',
      source: 'agent',
    }),
  (error: unknown) => error instanceof AppError && error.code === 'NON_TEXT_OBJECTS',
)

console.log('textProposal tests: PASS')
