import assert from 'node:assert/strict'

import {
  acceptPatchOperation,
  createDocumentPatchSet,
  getAcceptedPatchOperations,
  rejectPatchOperation,
  validatePatchSet,
} from '../src/utils/documentPatch.ts'

const patchSet = createDocumentPatchSet([
  {
    type: 'replace',
    targetText: '第一处',
    replacementText: '修改一',
    beforeContext: '标题：',
  },
  {
    type: 'delete',
    targetText: '第二处',
    replacementText: '',
  },
])

assert.equal(patchSet.operations.length, 2)
assert.equal(patchSet.status, 'pending')
assert.deepEqual(validatePatchSet(patchSet), [])

acceptPatchOperation(patchSet, patchSet.operations[0].id)
rejectPatchOperation(patchSet, patchSet.operations[1].id)
assert.deepEqual(
  getAcceptedPatchOperations(patchSet).map(operation => operation.id),
  [patchSet.operations[0].id],
)

const duplicateIds = structuredClone(patchSet)
duplicateIds.operations[1].id = duplicateIds.operations[0].id
assert.match(validatePatchSet(duplicateIds).join('\n'), /duplicate/i)

const invalidInsert = createDocumentPatchSet([{ type: 'insert_before', targetText: '', replacementText: 'x' }])
assert.match(validatePatchSet(invalidInsert).join('\n'), /targetText/i)

console.log('document patch tests: PASS')
