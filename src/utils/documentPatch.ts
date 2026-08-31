import { AppError } from '@/api/errors'
import { getProtectedObjects, hashText, TextChangeRisk } from '@/utils/textProposal'

export type PatchOperationType = 'replace' | 'delete' | 'insert_before' | 'insert_after'
export type PatchSetStatus =
  | 'pending'
  | 'applying'
  | 'applied'
  | 'cancelled'
  | 'conflicted'
  | 'rolled_back'
  | 'rollback_failed'
  | 'restored'
export type PatchOperationStatus = 'pending' | 'accepted' | 'rejected' | 'applied' | 'rolled_back' | 'conflicted'

export interface PatchOperationInput {
  type: PatchOperationType
  targetText: string
  replacementText: string
  beforeContext?: string
  afterContext?: string
  mapId?: string
  targetNodeId?: string
}

export interface DocumentPatchOperation extends PatchOperationInput {
  id: string
  beforeText: string
  beforeOoxml: string
  expectedBeforeHash: string
  expectedBeforeOoxmlHash: string
  anchorTag: string
  status: PatchOperationStatus
  risk: TextChangeRisk
  documentOrder: number
  appliedTextHash?: string
  appliedOoxmlHash?: string
}

export interface DocumentPatchSet {
  id: string
  operations: DocumentPatchOperation[]
  status: PatchSetStatus
  createdAt: string
}

const riskFor = (beforeText: string, replacementText: string): TextChangeRisk => {
  const ratio = beforeText.length === 0 ? 1 : Math.abs(replacementText.length - beforeText.length) / beforeText.length
  if (beforeText.length > 2000 || ratio > 0.75) return 'high'
  if (beforeText.length > 500 || ratio > 0.35) return 'medium'
  return 'low'
}

export const createDocumentPatchSet = (inputs: PatchOperationInput[]): DocumentPatchSet => ({
  id: crypto.randomUUID(),
  operations: inputs.map((input, index) => ({
    ...input,
    id: crypto.randomUUID(),
    beforeText: input.targetText,
    beforeOoxml: '',
    expectedBeforeHash: hashText(input.targetText),
    expectedBeforeOoxmlHash: '',
    anchorTag: '',
    status: 'pending',
    risk: riskFor(input.targetText, input.replacementText),
    documentOrder: index,
  })),
  status: 'pending',
  createdAt: new Date().toISOString(),
})

export const validatePatchSet = (patchSet: DocumentPatchSet): string[] => {
  const errors: string[] = []
  if (!patchSet.id) errors.push('PatchSet id is required')
  if (patchSet.operations.length === 0) errors.push('PatchSet must contain at least one operation')
  const ids = new Set<string>()
  patchSet.operations.forEach((operation, index) => {
    if (!operation.id) errors.push(`Operation ${index + 1} id is required`)
    if (ids.has(operation.id)) errors.push(`Duplicate operation id: ${operation.id}`)
    ids.add(operation.id)
    if (!operation.targetText) errors.push(`Operation ${index + 1} targetText is required`)
    if (!['replace', 'delete', 'insert_before', 'insert_after'].includes(operation.type))
      errors.push(`Operation ${index + 1} has an unsupported type`)
    if (operation.type === 'delete' && operation.replacementText !== '')
      errors.push(`Operation ${index + 1} delete replacementText must be empty`)
  })
  return errors
}

const findOperation = (patchSet: DocumentPatchSet, operationId: string): DocumentPatchOperation => {
  const operation = patchSet.operations.find(item => item.id === operationId)
  if (!operation) throw new AppError('REQUEST_FAILED', `Patch operation ${operationId} was not found`)
  return operation
}

export const acceptPatchOperation = (patchSet: DocumentPatchSet, operationId: string): DocumentPatchOperation => {
  const operation = findOperation(patchSet, operationId)
  if (operation.status !== 'pending' && operation.status !== 'rejected')
    throw new AppError('REQUEST_FAILED', 'Only pending patch operations can be accepted')
  operation.status = 'accepted'
  return operation
}

export const rejectPatchOperation = (patchSet: DocumentPatchSet, operationId: string): DocumentPatchOperation => {
  const operation = findOperation(patchSet, operationId)
  if (operation.status !== 'pending' && operation.status !== 'accepted')
    throw new AppError('REQUEST_FAILED', 'Only pending patch operations can be rejected')
  operation.status = 'rejected'
  return operation
}

export const getAcceptedPatchOperations = (patchSet: DocumentPatchSet): DocumentPatchOperation[] =>
  patchSet.operations.filter(operation => operation.status === 'accepted')

export const hydratePatchOperation = (
  operation: DocumentPatchOperation,
  beforeText: string,
  beforeOoxml: string,
  anchorTag: string,
): DocumentPatchOperation => ({
  ...operation,
  beforeText,
  beforeOoxml,
  expectedBeforeHash: hashText(beforeText),
  expectedBeforeOoxmlHash: hashText(beforeOoxml),
  anchorTag,
  risk: riskFor(beforeText, operation.replacementText),
  status: 'pending',
})

export const protectedObjectsInOperation = (operation: DocumentPatchOperation): string[] =>
  getProtectedObjects(operation.beforeOoxml)
