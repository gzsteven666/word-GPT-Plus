import assert from 'node:assert/strict'

import {
  createTaskTools,
  createTaskToolState,
  extractFormatChanges,
  isConfirmationIntent,
} from '../src/utils/taskTools.ts'

const readTools = createTaskTools(['read_document_structure', 'read_range'])
assert.deepEqual(
  readTools.map(tool => tool.name),
  ['read_document_structure', 'read_range'],
)

const editTools = createTaskTools([
  'read_range',
  'propose_document_patch',
  'apply_document_patch',
  'propose_format_patch',
  'apply_format_patch',
  'format_document_selection',
  'verify_patch',
])
assert.deepEqual(
  editTools.map(tool => tool.name),
  [
    'read_range',
    'propose_document_patch',
    'apply_document_patch',
    'propose_format_patch',
    'apply_format_patch',
    'format_document_selection',
    'verify_patch',
  ],
)

assert.equal(createTaskTools([]).length, 0)

// Shared state survives across createTaskTools calls (cross-turn proposals).
const sharedState = createTaskToolState()
const firstInstance = createTaskTools(['propose_document_patch'], undefined, sharedState)
const secondInstance = createTaskTools(['apply_document_patch'], undefined, sharedState)
assert.equal(firstInstance.length, 1)
assert.equal(secondInstance.length, 1)
assert.equal(firstInstance[0].name, 'propose_document_patch')
assert.equal(secondInstance[0].name, 'apply_document_patch')
assert.equal(sharedState.textProposals.size, 0)
assert.equal(sharedState.appliedTextChanges.size, 0)
assert.equal(sharedState.formatRequests.size, 0)
assert.equal(sharedState.activeFormatRequestId, null)
assert.equal(sharedState.appliedFormatRequests.size, 0)

// Fresh states are independent.
const freshState = createTaskToolState()
assert.notEqual(freshState.textProposals, sharedState.textProposals)

// extractFormatChanges drops control fields and undefined values.
assert.deepEqual(
  extractFormatChanges({
    scope: 'selection',
    operationId: 'op-1',
    formatId: 'fmt-1',
    fontSize: 12,
    lineSpacingMultiple: undefined,
    bold: true,
  }),
  { fontSize: 12, bold: true },
)
assert.deepEqual(extractFormatChanges({ scope: 'document' }), {})

// Confirmation intent detection.
for (const yes of [
  '是',
  '确定',
  '执行',
  '可以',
  '好的',
  '选中了',
  'ok',
  'Yes',
  'OK',
  'apply',
  'go ahead',
  '确认',
  '确定\n\n[Selected text: "Some selected text"]',
]) {
  assert.equal(isConfirmationIntent(yes), true, `expected confirmation: ${yes}`)
}
for (const no of ['请改写这段文字', '把第一段设置成12号字体', '查一下最新消息', '帮我总结一下', '', '  ']) {
  assert.equal(isConfirmationIntent(no), false, `expected non-confirmation: ${no}`)
}

console.log('taskTools tests: PASS')
