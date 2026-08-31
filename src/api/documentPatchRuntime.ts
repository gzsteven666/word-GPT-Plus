import { AppError } from '@/api/errors'
import {
  acceptPatchOperation,
  createDocumentPatchSet,
  DocumentPatchOperation,
  DocumentPatchSet,
  getAcceptedPatchOperations,
  hydratePatchOperation,
  PatchOperationInput,
} from '@/utils/documentPatch'
import { getProtectedObjects, hashText } from '@/utils/textProposal'

type RangeLike = Word.Range & {
  text?: string
  ooxml?: string
}

// Word can rewrite revision/identity metadata while a range is wrapped in a
// content control. Exclude those volatile attributes from the OOXML baseline
// while retaining the structural/text representation for conflict checks.
const canonicalizeOoxml = (ooxml: string): string =>
  (ooxml.match(/<w:body\b[\s\S]*?<\/w:body>/)?.[0] ?? ooxml)
    .replace(/<w:(?:proofErr|bookmarkStart|bookmarkEnd|lastRenderedPageBreak)\b[^>]*\/?\s*>/g, '')
    .replace(/<\/w:(?:proofErr|bookmarkStart|bookmarkEnd|lastRenderedPageBreak)>/g, '')
    .replace(/\s+w:(?:rsid[A-Za-z0-9]+|paraId|textId)="[^"]*"/g, '')
    .replace(/\s+w14:(?:paraId|textId)="[^"]*"/g, '')
    .replace(/\s+/g, ' ')
    .trim()

const hashOoxml = (ooxml: string): string => hashText(canonicalizeOoxml(ooxml))

const getControls = (context: Word.RequestContext) => {
  const controls = context.document.contentControls
  controls.load('items/tag')
  return controls
}

const findControl = (controls: Word.ContentControlCollection, tag: string): Word.ContentControl | undefined =>
  controls.items.find(item => item.tag === tag)

const rangeSnapshot = async (context: Word.RequestContext, range: RangeLike) => {
  const ooxml = range.getOoxml()
  range.load('text')
  await context.sync()
  return { text: range.text || '', ooxml: ooxml.value || '' }
}

const resolveCandidate = async (
  context: Word.RequestContext,
  input: PatchOperationInput,
): Promise<{ range: Word.Range; text: string; ooxml: string }> => {
  const results = context.document.body.search(input.targetText, { matchCase: true, matchWholeWord: false })
  results.load('items')
  await context.sync()
  const candidates = results.items as Word.Range[]
  if (candidates.length === 0) throw new AppError('DOCUMENT_CONFLICT', `Target text was not found: ${input.targetText}`)

  const matching: { range: Word.Range; text: string; ooxml: string }[] = []
  for (const candidate of candidates) {
    const snapshot = await rangeSnapshot(context, candidate as RangeLike)
    let contextText = snapshot.text
    const whole = typeof candidate.getRange === 'function' ? candidate.getRange('Whole') : undefined
    if (whole) {
      whole.load('text')
      await context.sync()
      contextText = whole.text || contextText
    }
    if (
      (!input.beforeContext || contextText.includes(`${input.beforeContext}${input.targetText}`)) &&
      (!input.afterContext || contextText.includes(`${input.targetText}${input.afterContext}`))
    )
      matching.push({ range: candidate, ...snapshot })
  }
  if (matching.length !== 1)
    throw new AppError(
      'DOCUMENT_CONFLICT',
      matching.length === 0
        ? `Target context was not found: ${input.targetText}`
        : `Target text is ambiguous: ${input.targetText}`,
    )
  return matching[0]
}

const rejectOverlaps = async (context: Word.RequestContext, ranges: Word.Range[]): Promise<void> => {
  for (let left = 0; left < ranges.length; left += 1) {
    for (let right = left + 1; right < ranges.length; right += 1) {
      if (typeof ranges[left].compareLocationWith !== 'function')
        throw new AppError('WORD_API_UNSUPPORTED', 'Word cannot compare patch range locations')
      const relation = ranges[left].compareLocationWith(ranges[right])
      await context.sync()
      if (['Equal', 'Overlap', 'Inside', 'Contains'].includes(String(relation.value)))
        throw new AppError('DOCUMENT_CONFLICT', 'Patch operations overlap and cannot be applied together')
    }
  }
}

export const proposeDocumentPatchSet = async (inputs: PatchOperationInput[]): Promise<DocumentPatchSet> =>
  Word.run(async context => {
    const patchSet = createDocumentPatchSet(inputs)
    const resolved: { range: Word.Range; operation: DocumentPatchOperation; control?: Word.ContentControl }[] = []
    for (const operation of patchSet.operations) {
      const candidate = await resolveCandidate(context, operation)
      const protectedObjects = getProtectedObjects(candidate.ooxml)
      if (protectedObjects.length > 0)
        throw new AppError('NON_TEXT_OBJECTS', `Target contains protected objects: ${protectedObjects.join(', ')}`)
      resolved.push({
        range: candidate.range,
        operation: hydratePatchOperation(
          operation,
          candidate.text,
          candidate.ooxml,
          `wordgpt_patch_pending_${operation.id.replaceAll('-', '')}`,
        ),
      })
    }
    await rejectOverlaps(
      context,
      resolved.map(item => item.range),
    )
    try {
      for (const item of resolved) {
        const control = item.range.insertContentControl('RichText')
        control.tag = item.operation.anchorTag
        control.title = 'Word GPT pending patch'
        control.appearance = 'Hidden'
        item.control = control
      }
      await context.sync()

      // Word may serialize the content differently once it is wrapped by a
      // hidden content control. Baseline the anchored content itself so the
      // preflight hash checks detect real edits without rejecting the anchor.
      for (const item of resolved) {
        if (!item.control) throw new AppError('WORD_API_UNSUPPORTED', 'Word did not return a patch anchor')
        const anchored = await rangeSnapshot(context, item.control.getRange('Content') as RangeLike)
        item.operation.expectedBeforeHash = hashText(anchored.text)
        item.operation.expectedBeforeOoxmlHash = hashOoxml(anchored.ooxml)
      }
    } catch (error) {
      for (const item of resolved) {
        const controls = getControls(context)
        await context.sync()
        const control = findControl(controls, item.operation.anchorTag)
        control?.delete(true)
      }
      await context.sync()
      throw error
    }
    patchSet.operations = resolved.map(item => item.operation)
    return patchSet
  })

const assertOperationCurrent = async (context: Word.RequestContext, operation: DocumentPatchOperation) => {
  const controls = getControls(context)
  await context.sync()
  const control = findControl(controls, operation.anchorTag)
  if (!control) throw new AppError('DOCUMENT_CONFLICT', 'The patch anchor is no longer available')
  const range = control.getRange('Content') as RangeLike
  const snapshot = await rangeSnapshot(context, range)
  if (hashText(snapshot.text) !== operation.expectedBeforeHash)
    throw new AppError('DOCUMENT_CONFLICT', 'A patch target changed before application')
  if (hashOoxml(snapshot.ooxml) !== operation.expectedBeforeOoxmlHash)
    throw new AppError('DOCUMENT_CONFLICT', 'A patch target formatting changed before application')
  const protectedObjects = getProtectedObjects(snapshot.ooxml)
  if (protectedObjects.length > 0)
    throw new AppError('NON_TEXT_OBJECTS', `Target contains protected objects: ${protectedObjects.join(', ')}`)
  return { control, range }
}

const applyOperation = async (operation: DocumentPatchOperation): Promise<void> =>
  Word.run(async context => {
    const { control, range } = await assertOperationCurrent(context, operation)
    const location =
      operation.type === 'replace' || operation.type === 'delete'
        ? 'Replace'
        : operation.type === 'insert_before'
          ? 'Before'
          : 'After'
    range.insertText(operation.replacementText, location)
    const appliedOoxml = range.getOoxml()
    range.load('text')
    await context.sync()
    operation.appliedTextHash = hashText(range.text || '')
    operation.appliedOoxmlHash = hashOoxml(appliedOoxml.value || '')
    control.title = 'Word GPT applied patch'
  })

const restoreOperation = async (operation: DocumentPatchOperation): Promise<void> =>
  Word.run(async context => {
    const { control, range } = await assertAppliedCurrent(context, operation)
    range.insertOoxml(operation.beforeOoxml, 'Replace')
    await context.sync()
    control.delete(true)
    await context.sync()
  })

const assertAppliedCurrent = async (context: Word.RequestContext, operation: DocumentPatchOperation) => {
  const controls = getControls(context)
  await context.sync()
  const control = findControl(controls, operation.anchorTag)
  if (!control) throw new AppError('DOCUMENT_CONFLICT', 'The applied patch anchor is no longer available')
  const range = control.getRange('Content') as RangeLike
  const snapshot = await rangeSnapshot(context, range)
  if (operation.appliedTextHash && hashText(snapshot.text) !== operation.appliedTextHash)
    throw new AppError('DOCUMENT_CONFLICT', 'The applied patch text changed after application')
  if (operation.appliedOoxmlHash && hashOoxml(snapshot.ooxml) !== operation.appliedOoxmlHash)
    throw new AppError('DOCUMENT_CONFLICT', 'The applied patch formatting changed after application')
  return { control, range }
}

export const applyDocumentPatchSet = async (patchSet: DocumentPatchSet): Promise<DocumentPatchSet> => {
  const accepted = getAcceptedPatchOperations(patchSet)
  if (accepted.length === 0) throw new AppError('REQUEST_FAILED', 'No patch operations were accepted')
  try {
    await Word.run(async context => {
      for (const operation of accepted) await assertOperationCurrent(context, operation)
    })
  } catch (error) {
    patchSet.status = 'conflicted'
    await releaseDocumentPatchSetAnchors(patchSet)
    throw error
  }
  patchSet.status = 'applying'
  const applied: DocumentPatchOperation[] = []
  try {
    for (const operation of [...accepted].sort((a, b) => b.documentOrder - a.documentOrder)) {
      await applyOperation(operation)
      operation.status = 'applied'
      applied.push(operation)
    }
    patchSet.status = 'applied'
    return patchSet
  } catch (_error) {
    let rollbackFailed = false
    for (const operation of [...applied].reverse()) {
      try {
        await restoreOperation(operation)
        operation.status = 'rolled_back'
      } catch {
        rollbackFailed = true
      }
    }
    patchSet.status = rollbackFailed ? 'rollback_failed' : 'rolled_back'
    if (!rollbackFailed) await releaseDocumentPatchSetAnchors(patchSet)
    throw new AppError(
      'REQUEST_FAILED',
      rollbackFailed ? 'Patch failed and rollback failed' : 'Patch failed and was rolled back',
    )
  }
}

export const verifyDocumentPatchSet = async (patchSet: DocumentPatchSet) => {
  for (const operation of patchSet.operations.filter(item => item.status === 'applied')) {
    await Word.run(async context => {
      await assertAppliedCurrent(context, operation)
    })
  }
  return { status: 'verified' as const, patchSetId: patchSet.id }
}

export const restoreDocumentPatchSet = async (patchSet: DocumentPatchSet): Promise<DocumentPatchSet> => {
  const applied = patchSet.operations.filter(operation => operation.status === 'applied')
  await Word.run(async context => {
    for (const operation of applied) await assertAppliedCurrent(context, operation)
  })
  for (const operation of applied) await restoreOperation(operation)
  applied.forEach(operation => {
    operation.status = 'rolled_back'
  })
  patchSet.status = 'restored'
  return patchSet
}

export const releaseDocumentPatchSetAnchors = async (patchSet: DocumentPatchSet): Promise<void> =>
  Word.run(async context => {
    const controls = getControls(context)
    await context.sync()
    patchSet.operations.forEach(operation => findControl(controls, operation.anchorTag)?.delete(true))
    await context.sync()
  })

export { acceptPatchOperation }
