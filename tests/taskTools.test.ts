import assert from 'node:assert/strict'

import { createTaskTools } from '../src/utils/taskTools.ts'

const readTools = createTaskTools(['read_document_structure', 'read_range'])
assert.deepEqual(
  readTools.map(tool => tool.name),
  ['read_document_structure', 'read_range'],
)

const editTools = createTaskTools([
  'read_range',
  'propose_document_patch',
  'apply_document_patch',
  'format_document_selection',
  'verify_patch',
])
assert.deepEqual(
  editTools.map(tool => tool.name),
  ['read_range', 'propose_document_patch', 'apply_document_patch', 'format_document_selection', 'verify_patch'],
)

assert.equal(createTaskTools([]).length, 0)
console.log('taskTools tests: PASS')
