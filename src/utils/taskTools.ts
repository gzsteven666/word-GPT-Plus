import { tool } from '@langchain/core/tools'
import { z } from 'zod'

import { createDocumentMap, readDocumentMapNodes } from '@/api/documentMapRuntime'
import {
  applyDocumentPatchSet,
  proposeDocumentPatchSet,
  releaseDocumentPatchSetAnchors,
  verifyDocumentPatchSet,
} from '@/api/documentPatchRuntime'
import { AppError } from '@/api/errors'
import {
  AppliedTextChange,
  applyTextChangeProposal,
  createSelectionAnchor,
  getAnchoredRange,
  makeSelectionInsertProposal,
  makeSelectionProposal,
  readSelectionInspection,
  readSelectionTextSnapshot,
  releaseSelectionAnchor,
  restoreTextChange,
  verifyTextChange,
} from '@/api/safeEdit'
import { DocumentMap, queryDocumentMap } from '@/utils/documentMap'
import { DocumentPatchSet, PatchOperationInput } from '@/utils/documentPatch'
import { buildTextDiff, hashText, TextChangeProposal } from '@/utils/textProposal'

export type TaskToolName =
  | 'read_document_structure'
  | 'create_document_map'
  | 'query_document_map'
  | 'read_document_nodes'
  | 'read_range'
  | 'propose_document_patch'
  | 'apply_document_patch'
  | 'propose_document_patch_set'
  | 'apply_document_patch_set'
  | 'verify_document_patch_set'
  | 'propose_format_patch'
  | 'apply_format_patch'
  | 'format_document_selection'
  | 'verify_patch'

export interface FormatRequest {
  id: string
  scope: 'selection' | 'document'
  changes: Record<string, boolean | number | string>
  beforeHash?: string
  beforeOoxml?: string
  beforeOoxmlHash?: string
  anchorTag?: string
  appliedOoxmlHash?: string
}

/**
 * Shared state passed between `createTaskTools` calls so proposals and applied
 * results survive across agent turns (the tool list is rebuilt for every
 * user message). Without this, a proposal created in one turn is gone by the
 * time the user confirms in the next turn.
 */
export interface TaskToolState {
  textProposals: Map<string, TextChangeProposal>
  appliedTextChanges: Map<string, AppliedTextChange>
  formatRequests: Map<string, FormatRequest>
  activeFormatRequestId: string | null
  appliedFormatRequests: Map<string, FormatRequest>
  patchSets: Map<string, DocumentPatchSet>
  documentMaps: Map<string, DocumentMap>
}

export const createTaskToolState = (): TaskToolState => ({
  textProposals: new Map(),
  appliedTextChanges: new Map(),
  formatRequests: new Map(),
  activeFormatRequestId: null,
  appliedFormatRequests: new Map(),
  patchSets: new Map(),
  documentMaps: new Map(),
})

export interface TaskToolSecurity {
  requestTextChangeApproval?: (proposal: TextChangeProposal) => Promise<boolean>
  requestSensitiveDataApproval?: (scope: string) => Promise<boolean>
  requestPatchSetApproval?: (patchSet: DocumentPatchSet) => Promise<boolean>
  onTextChangeApplied?: (change: AppliedTextChange) => void
  onPatchSetApplied?: (patchSet: DocumentPatchSet) => void
  onFormatApplied?: (request: FormatRequest) => void
  allowedFormatFields?: string[]
}

export const extractUserInstruction = (text: string): string => {
  const selectionMarker = '\n\n[Selected text: "'
  const markerIndex = text.indexOf(selectionMarker)
  return (markerIndex >= 0 ? text.slice(0, markerIndex) : text).trim()
}

export const isFormatPreviewIntent = (text: string): boolean => {
  const instruction = extractUserInstruction(text)
  return /先.{0,20}(?:方案|预览)|(?:等待|等我).{0,12}(?:确认|同意)|确认后再|(?:show|display).{0,20}(?:plan|preview).{0,20}(?:first|before)|wait.{0,12}(?:confirm|approval)|preview.{0,20}before.{0,12}(?:apply|execut)/i.test(
    instruction,
  )
}

/**
 * Detect a short confirmation/continuation reply. This only controls which
 * tool is made available or required; it never performs the write itself.
 */
export const isConfirmationIntent = (text: string): boolean =>
  /^(?:(?:是|好|好的|确定|确认|可以|执行|应用|继续|同意|没问题|选中了|已选中|完成)(?:吧|了|啦|即可|就行|好了)?|就按(?:这个|该)?方案(?:来|执行)?|按这个来|照此执行|yes|ok|okay|confirm|apply|proceed|go ahead|sure|selected|done|accept|fine)[！!。.，,\s]*$/i.test(
    extractUserInstruction(text),
  )

export type FormatToolRoute = 'propose_format_patch' | 'apply_format_patch' | 'format_document_selection'

export const inferRequestedFormatFields = (text: string): string[] => {
  const instruction = extractUserInstruction(text)
  const fields = new Set<string>()
  if (/字号|号字体|\d+(?:\.\d+)?\s*(?:磅|pt\b)|font\s*size/i.test(instruction)) fields.add('fontSize')
  if (
    /宋体|微软雅黑|黑体|仿宋|楷体|calibri|arial|times new roman|font\s*family|字体.{0,6}(?:改|设|设置|更换)?为(?!\s*\d)/i.test(
      instruction,
    )
  )
    fields.add('fontName')
  if (/粗体|加粗|\bbold\b/i.test(instruction)) fields.add('bold')
  if (/斜体|\bitalic/i.test(instruction)) fields.add('italic')
  if (/下划线|underline/i.test(instruction)) fields.add('underline')
  if (/字体颜色|字色|font\s*color/i.test(instruction)) fields.add('fontColor')
  if (/高亮|highlight/i.test(instruction)) fields.add('highlightColor')
  if (/对齐|align/i.test(instruction)) fields.add('alignment')
  if (/行距|line\s*spacing/i.test(instruction)) {
    fields.add('lineSpacing')
    fields.add('lineSpacingMultiple')
  }
  if (/段后|space\s*after/i.test(instruction)) fields.add('spaceAfter')
  return [...fields]
}

export const resolveFormatToolRoute = (text: string, hasActiveProposal: boolean): FormatToolRoute | null => {
  const instruction = extractUserInstruction(text)
  if (hasActiveProposal && isConfirmationIntent(instruction)) return 'apply_format_patch'
  if (!/格式|排版|美化|字体|字号|行距|对齐|style|format|layout|beautify/i.test(instruction)) return null
  return isFormatPreviewIntent(instruction) ? 'propose_format_patch' : 'format_document_selection'
}

const proposalSummary = (proposal: TextChangeProposal) =>
  JSON.stringify({
    proposalId: proposal.id,
    operation: proposal.operation,
    scope: proposal.scope,
    beforeText: proposal.beforeText,
    afterText: proposal.afterText,
    expectedBeforeHash: proposal.expectedBeforeHash,
    risk: proposal.risk,
    status: proposal.status,
    diff: buildTextDiff(proposal.beforeText, proposal.afterText),
  })

const createProposalFromSelection = async (
  operation: TextChangeProposal['operation'],
  afterText: string,
): Promise<TextChangeProposal> => {
  if (operation === 'insert') return makeSelectionInsertProposal(afterText, 'End', 'agent')
  return makeSelectionProposal(afterText, 'agent', operation)
}

const readDocumentHash = async (): Promise<string> =>
  Word.run(async context => {
    const body = context.document.body
    body.load('text')
    await context.sync()
    return hashText(body.text || '')
  })

const readCurrentHash = async (scope: FormatRequest['scope']): Promise<string> =>
  scope === 'selection' ? (await readSelectionTextSnapshot()).hash : readDocumentHash()

const readDocumentOoxml = async (): Promise<string> =>
  Word.run(async context => {
    const range = context.document.body.getRange()
    const ooxml = range.getOoxml()
    await context.sync()
    return ooxml.value || ''
  })

const readFormatOoxmlHash = async (request: FormatRequest): Promise<string> =>
  Word.run(async context => {
    const controls = request.anchorTag ? context.document.contentControls : undefined
    controls?.load('items/tag')
    if (controls) await context.sync()
    const range = request.anchorTag
      ? getAnchoredRange(controls!, request.anchorTag)
      : request.scope === 'document'
        ? context.document.body.getRange()
        : context.document.getSelection()
    if (!range) throw new AppError('DOCUMENT_CONFLICT', 'The formatting anchor is no longer available')
    const ooxml = range.getOoxml()
    await context.sync()
    return hashText(ooxml.value || '')
  })

/**
 * Extract the formatting-change map from tool arguments, dropping control
 * fields (scope, operationId, formatId) and undefined values.
 */
export const extractFormatChanges = (args: Record<string, unknown>): Record<string, boolean | number | string> =>
  Object.fromEntries(
    Object.entries(args).filter(
      ([key, value]) => !['scope', 'operationId', 'formatId'].includes(key) && value !== undefined,
    ),
  ) as Record<string, boolean | number | string>

const restrictFormatChanges = (
  changes: Record<string, boolean | number | string>,
  allowedFields?: string[],
): Record<string, boolean | number | string> => {
  if (!allowedFields?.length) return changes
  const allowed = new Set(allowedFields)
  return Object.fromEntries(Object.entries(changes).filter(([key]) => allowed.has(key)))
}

const applyFormatRequest = async (request: FormatRequest): Promise<void> => {
  const changes = request.changes
  await Word.run(async context => {
    const controls = request.anchorTag ? context.document.contentControls : undefined
    controls?.load('items/tag')
    if (controls) await context.sync()
    const range = request.anchorTag
      ? getAnchoredRange(controls!, request.anchorTag)
      : request.scope === 'document'
        ? context.document.body.getRange()
        : context.document.getSelection()
    if (!range) throw new AppError('DOCUMENT_CONFLICT', 'The formatting anchor is no longer available')
    const currentOoxml = range.getOoxml()
    range.load('text')
    await context.sync()
    if (request.beforeHash && hashText(range.text || '') !== request.beforeHash) {
      throw new AppError('DOCUMENT_CONFLICT', 'The selection changed while formatting was awaiting approval')
    }
    if (request.beforeOoxmlHash && hashText(currentOoxml.value || '') !== request.beforeOoxmlHash) {
      throw new AppError('DOCUMENT_CONFLICT', 'The selection formatting changed while approval was pending')
    }

    if (changes.bold !== undefined) range.font.bold = changes.bold as boolean
    if (changes.italic !== undefined) range.font.italic = changes.italic as boolean
    if (changes.underline !== undefined) range.font.underline = changes.underline ? 'Single' : 'None'
    if (changes.fontName) range.font.name = changes.fontName as string
    if (changes.fontSize !== undefined) range.font.size = changes.fontSize as number
    if (changes.fontColor) range.font.color = changes.fontColor as string
    if (changes.highlightColor) range.font.highlightColor = changes.highlightColor as string

    const hasParagraphChanges =
      changes.alignment !== undefined ||
      changes.lineSpacing !== undefined ||
      changes.lineSpacingMultiple !== undefined ||
      changes.spaceAfter !== undefined
    if (!hasParagraphChanges) {
      await context.sync()
      return
    }

    // For relative line spacing we need the effective font size (1 line = font size in points).
    let fontSize = changes.fontSize !== undefined ? (changes.fontSize as number) : 12
    if (changes.fontSize === undefined) {
      range.font.load('size')
      await context.sync()
      const currentSize = range.font.size
      if (typeof currentSize === 'number' && currentSize > 0) fontSize = currentSize
    }

    const paragraphs = range.paragraphs
    paragraphs.load('items')
    await context.sync()
    for (const paragraph of paragraphs.items) {
      if (changes.alignment) paragraph.alignment = changes.alignment as Word.Alignment
      if (changes.lineSpacingMultiple !== undefined) {
        paragraph.lineSpacing = (changes.lineSpacingMultiple as number) * fontSize
      } else if (changes.lineSpacing !== undefined) {
        paragraph.lineSpacing = changes.lineSpacing as number
      }
      if (changes.spaceAfter !== undefined) paragraph.spaceAfter = changes.spaceAfter as number
    }
    await context.sync()
  })
}

export interface FormatVerification {
  verified: boolean
  checks: {
    key: string
    expected: string | number | boolean
    actual: string | number | boolean | null
    ok: boolean
  }[]
}

const verifyFormatRequest = async (request: FormatRequest): Promise<FormatVerification> =>
  Word.run(async context => {
    const changes = request.changes
    const controls = request.anchorTag ? context.document.contentControls : undefined
    controls?.load('items/tag')
    if (controls) await context.sync()
    const range = request.anchorTag
      ? getAnchoredRange(controls!, request.anchorTag)
      : request.scope === 'document'
        ? context.document.body.getRange()
        : context.document.getSelection()
    if (!range) throw new AppError('DOCUMENT_CONFLICT', 'The formatting anchor is no longer available')
    range.font.load('name,size,bold,italic,underline,color,highlightColor')
    const paragraphs = range.paragraphs
    paragraphs.load('items')
    await context.sync()
    paragraphs.items.forEach(paragraph => paragraph.load('alignment,lineSpacing,spaceAfter'))
    await context.sync()

    const checks: FormatVerification['checks'] = []
    const addCheck = (key: string, expected: unknown, actual: unknown, ok: boolean) => {
      checks.push({
        key,
        expected: String(expected),
        actual: actual === null || actual === undefined ? null : String(actual),
        ok,
      })
    }
    const eqNum = (a: unknown, b: unknown) => typeof a === 'number' && typeof b === 'number' && Math.abs(a - b) < 0.5
    const eqStr = (a: unknown, b: unknown) =>
      typeof a === 'string' && typeof b === 'string' && a.toLowerCase() === b.toLowerCase()
    const eqBool = (a: unknown, b: unknown) => typeof a === 'boolean' && a === b

    if (changes.bold !== undefined)
      addCheck('bold', changes.bold, range.font.bold, eqBool(range.font.bold, changes.bold))
    if (changes.italic !== undefined)
      addCheck('italic', changes.italic, range.font.italic, eqBool(range.font.italic, changes.italic))
    if (changes.underline !== undefined) {
      const expectedUnderline = changes.underline ? 'Single' : 'None'
      const actualUnderline =
        typeof range.font.underline === 'string' ? (range.font.underline as string) : String(range.font.underline)
      addCheck('underline', expectedUnderline, actualUnderline, actualUnderline === expectedUnderline)
    }
    if (changes.fontName)
      addCheck('fontName', changes.fontName, range.font.name, eqStr(changes.fontName, range.font.name))
    if (changes.fontSize !== undefined)
      addCheck('fontSize', changes.fontSize, range.font.size, eqNum(changes.fontSize, range.font.size))
    if (changes.fontColor)
      addCheck('fontColor', changes.fontColor, range.font.color, eqStr(changes.fontColor, range.font.color))
    if (changes.highlightColor)
      addCheck(
        'highlightColor',
        changes.highlightColor,
        range.font.highlightColor,
        eqStr(changes.highlightColor, range.font.highlightColor),
      )

    const first = paragraphs.items[0]
    if (first) {
      if (changes.alignment)
        addCheck('alignment', changes.alignment, first.alignment, first.alignment === changes.alignment)
      if (changes.lineSpacingMultiple !== undefined) {
        const fontSize =
          changes.fontSize !== undefined
            ? (changes.fontSize as number)
            : typeof range.font.size === 'number'
              ? range.font.size
              : 12
        const expectedPoints = (changes.lineSpacingMultiple as number) * fontSize
        addCheck(
          'lineSpacing',
          `${changes.lineSpacingMultiple}x`,
          first.lineSpacing,
          eqNum(expectedPoints, first.lineSpacing),
        )
      } else if (changes.lineSpacing !== undefined) {
        addCheck('lineSpacing', changes.lineSpacing, first.lineSpacing, eqNum(changes.lineSpacing, first.lineSpacing))
      }
      if (changes.spaceAfter !== undefined)
        addCheck('spaceAfter', changes.spaceAfter, first.spaceAfter, eqNum(changes.spaceAfter, first.spaceAfter))
    }

    return { verified: checks.length > 0 && checks.every(check => check.ok), checks }
  })

export const createTaskTools = (
  enabledTools?: TaskToolName[],
  security?: TaskToolSecurity,
  state: TaskToolState = createTaskToolState(),
) => {
  const shouldEnable = (name: TaskToolName) => !enabledTools || enabledTools.includes(name)
  const tools = []

  if (shouldEnable('create_document_map')) {
    tools.push(
      tool(
        async () => {
          if (security?.requestSensitiveDataApproval) {
            const approved = await security.requestSensitiveDataApproval('document headings and paragraph previews')
            if (!approved) return 'The user did not authorize sharing the document map with the model.'
          }
          const map = await createDocumentMap()
          state.documentMaps.set(map.id, map)
          const headingNodes = map.nodes.filter(node => node.kind === 'heading')
          const headings = headingNodes
            .slice(0, 100)
            .map(node => ({ nodeId: node.id, level: node.headingLevel, path: node.headingPath, preview: node.preview }))
          return JSON.stringify({
            mapId: map.id,
            nodeCount: map.nodes.length,
            headingCount: headingNodes.length,
            protectedObjectsAvailable: map.nodes.every(node => node.protectedObjectsAvailable),
            headings,
          })
        },
        {
          name: 'create_document_map',
          description:
            'Build a read-only, in-memory map of Word headings and paragraphs. It returns node IDs, heading paths, counts, and bounded previews; it never changes the document.',
          schema: z.object({}),
        },
      ),
    )
  }

  if (shouldEnable('query_document_map')) {
    tools.push(
      tool(
        async input => {
          const args = input as { mapId: string; query: string; limit?: number }
          const map = state.documentMaps.get(args.mapId)
          if (!map) throw new AppError('REQUEST_FAILED', 'The document map was not found or has expired')
          return JSON.stringify(
            queryDocumentMap(map, args.query, args.limit).map(result => ({
              nodeId: result.node.id,
              score: result.score,
              kind: result.node.kind,
              order: result.node.order,
              headingPath: result.node.headingPath,
              preview: result.node.preview,
              protectedObjectsAvailable: result.node.protectedObjectsAvailable,
            })),
          )
        },
        {
          name: 'query_document_map',
          description:
            'Search an existing document map by heading path and paragraph keywords. Returns ranked node IDs and short previews without reading the full document.',
          schema: z.object({
            mapId: z.string(),
            query: z.string().min(1),
            limit: z.number().int().positive().max(50).optional(),
          }),
        },
      ),
    )
  }

  if (shouldEnable('read_document_nodes')) {
    tools.push(
      tool(
        async input => {
          const args = input as { mapId: string; nodeIds: string[]; maxChars?: number }
          const map = state.documentMaps.get(args.mapId)
          if (!map) throw new AppError('REQUEST_FAILED', 'The document map was not found or has expired')
          if (security?.requestSensitiveDataApproval) {
            const approved = await security.requestSensitiveDataApproval('selected document nodes')
            if (!approved) return 'The user did not authorize sharing the selected document nodes.'
          }
          return JSON.stringify(await readDocumentMapNodes(map, args.nodeIds, args.maxChars))
        },
        {
          name: 'read_document_nodes',
          description:
            'Read only the selected nodes from a document map, bounded by maxChars. Rechecks the map hashes and rejects stale or moved nodes.',
          schema: z.object({
            mapId: z.string(),
            nodeIds: z.array(z.string()).min(1).max(50),
            maxChars: z.number().int().positive().max(50000).optional(),
          }),
        },
      ),
    )
  }

  if (shouldEnable('read_document_structure')) {
    tools.push(
      tool(
        async () => {
          if (security?.requestSensitiveDataApproval) {
            const approved = await security.requestSensitiveDataApproval('document structure and paragraph text')
            if (!approved) return 'The user did not authorize sharing document structure with the model.'
          }
          return Word.run(async context => {
            const paragraphs = context.document.body.paragraphs
            paragraphs.load('items/text')
            await context.sync()
            const items = paragraphs.items
              .slice(0, 200)
              .map((paragraph, index) => `${index + 1}. ${paragraph.text || ''}`)
            return JSON.stringify({ paragraphCount: paragraphs.items.length, paragraphs: items })
          })
        },
        {
          name: 'read_document_structure',
          description:
            'Read a compact, read-only outline of the Word document. Use this before reviewing a section or locating document structure; it does not return the full document body.',
          schema: z.object({}),
        },
      ),
    )
  }

  if (shouldEnable('read_range')) {
    tools.push(
      tool(
        async () => {
          const snapshot = await readSelectionInspection()
          return JSON.stringify({
            text: snapshot.text,
            hash: snapshot.hash,
            protectedObjects: snapshot.protectedObjects,
            protectedObjectsAvailable: snapshot.protectedObjectsAvailable,
            warning: snapshot.warning,
          })
        },
        {
          name: 'read_range',
          description:
            'Read the currently selected Word range and content hash. Protected-object markers are included when Word can provide OOXML; check protectedObjectsAvailable before relying on them. This is read-only.',
          schema: z.object({}),
        },
      ),
    )
  }

  if (shouldEnable('propose_document_patch')) {
    tools.push(
      tool(
        async input => {
          const args = input as { operation: TextChangeProposal['operation']; afterText: string }
          const proposal = await createProposalFromSelection(args.operation, args.afterText)
          state.textProposals.set(proposal.id, proposal)
          return proposalSummary(proposal)
        },
        {
          name: 'propose_document_patch',
          description:
            'Create a reviewable text patch for the current Word selection. This does not alter visible content, but creates a hidden temporary anchor so the original range remains stable. Call apply_document_patch only after the proposal is ready for user review.',
          schema: z.object({
            operation: z.enum(['replace', 'insert', 'delete']),
            afterText: z
              .string()
              .describe('The proposed replacement or inserted text; use an empty string for delete.'),
          }),
        },
      ),
    )
  }

  if (shouldEnable('apply_document_patch')) {
    tools.push(
      tool(
        async input => {
          const args = input as { proposalId: string }
          const existing = state.appliedTextChanges.get(args.proposalId)
          if (existing) {
            return JSON.stringify({ proposalId: args.proposalId, status: 'already_applied', idempotent: true })
          }

          const proposal = state.textProposals.get(args.proposalId)
          if (!proposal) throw new AppError('REQUEST_FAILED', 'The patch proposal was not found or has expired')
          if (!security?.requestTextChangeApproval) {
            throw new AppError('REQUEST_FAILED', 'A user approval handler is required before applying a document patch')
          }
          if (!(await security.requestTextChangeApproval(proposal))) {
            return JSON.stringify({ proposalId: proposal.id, status: 'cancelled' })
          }

          const applied = await applyTextChangeProposal(proposal)
          state.appliedTextChanges.set(proposal.id, applied)
          security.onTextChangeApplied?.(applied)
          return JSON.stringify({ proposalId: proposal.id, status: 'applied', idempotent: false })
        },
        {
          name: 'apply_document_patch',
          description:
            'Apply a previously proposed patch only after explicit user approval. The operation is guarded by the original selection hash and proposal ID.',
          schema: z.object({ proposalId: z.string().describe('The proposalId returned by propose_document_patch') }),
        },
      ),
    )
  }

  if (shouldEnable('propose_document_patch_set')) {
    tools.push(
      tool(
        async input => {
          const args = input as { operations: PatchOperationInput[] }
          const mapIds = [...new Set(args.operations.map(operation => operation.mapId).filter(Boolean))]
          if (mapIds.length > 1) throw new AppError('DOCUMENT_CONFLICT', 'A PatchSet cannot use multiple document maps')
          const documentMap = mapIds[0] ? state.documentMaps.get(mapIds[0]) : undefined
          if (mapIds[0] && !documentMap)
            throw new AppError('REQUEST_FAILED', 'The document map was not found or has expired')
          const patchSet = await proposeDocumentPatchSet(args.operations, documentMap)
          state.patchSets.set(patchSet.id, patchSet)
          return JSON.stringify({
            patchSetId: patchSet.id,
            status: patchSet.status,
            operations: patchSet.operations.map(operation => ({
              operationId: operation.id,
              type: operation.type,
              beforeText: operation.beforeText,
              afterText: operation.replacementText,
              risk: operation.risk,
              status: operation.status,
              diff: buildTextDiff(operation.beforeText, operation.replacementText),
            })),
          })
        },
        {
          name: 'propose_document_patch_set',
          description:
            'Create a multi-location text PatchSet for Word. Each target must be uniquely identified by targetText and optional beforeContext/afterContext. This creates hidden anchors but does not change visible text. Present every operation for review before calling apply_document_patch_set.',
          schema: z.object({
            operations: z
              .array(
                z.object({
                  type: z.enum(['replace', 'delete', 'insert_before', 'insert_after']),
                  targetText: z.string().min(1),
                  replacementText: z.string(),
                  beforeContext: z.string().optional(),
                  afterContext: z.string().optional(),
                  mapId: z.string().optional(),
                  targetNodeId: z.string().optional(),
                }),
              )
              .min(1),
          }),
        },
      ),
    )
  }

  if (shouldEnable('apply_document_patch_set')) {
    tools.push(
      tool(
        async input => {
          const args = input as { patchSetId: string }
          const patchSet = state.patchSets.get(args.patchSetId)
          if (!patchSet) throw new AppError('REQUEST_FAILED', 'The PatchSet was not found or has expired')
          if (patchSet.status === 'applied')
            return JSON.stringify({ patchSetId: patchSet.id, status: 'already_applied', idempotent: true })
          if (!security?.requestPatchSetApproval)
            throw new AppError('REQUEST_FAILED', 'A user approval handler is required before applying a PatchSet')
          if (!(await security.requestPatchSetApproval(patchSet))) {
            patchSet.status = 'cancelled'
            await releaseDocumentPatchSetAnchors(patchSet)
            return JSON.stringify({ patchSetId: patchSet.id, status: 'cancelled' })
          }
          const applied = await applyDocumentPatchSet(patchSet)
          state.patchSets.set(applied.id, applied)
          security.onPatchSetApplied?.(applied)
          return JSON.stringify({ patchSetId: applied.id, status: applied.status, idempotent: false })
        },
        {
          name: 'apply_document_patch_set',
          description:
            'Apply only the accepted operations in a previously proposed PatchSet after explicit per-item user review. All accepted targets are preflight-checked before any write; failures trigger compensating rollback.',
          schema: z.object({ patchSetId: z.string() }),
        },
      ),
    )
  }

  if (shouldEnable('verify_document_patch_set')) {
    tools.push(
      tool(
        async input => {
          const args = input as { patchSetId: string }
          const patchSet = state.patchSets.get(args.patchSetId)
          if (!patchSet) throw new AppError('REQUEST_FAILED', 'The PatchSet was not found or has expired')
          return JSON.stringify(await verifyDocumentPatchSet(patchSet))
        },
        {
          name: 'verify_document_patch_set',
          description:
            'Verify that every applied operation in a PatchSet still matches its applied text and OOXML hashes.',
          schema: z.object({ patchSetId: z.string() }),
        },
      ),
    )
  }

  const formatFieldsSchema = {
    scope: z.enum(['selection', 'document']).default('selection'),
    bold: z.boolean().optional().describe('Bold (粗体).'),
    italic: z.boolean().optional().describe('Italic (斜体).'),
    underline: z.boolean().optional().describe('Underline (下划线).'),
    fontName: z.string().optional().describe('Font family name, e.g. "Calibri" or "宋体".'),
    fontSize: z.number().positive().max(200).optional().describe('Font size in points (字号，磅). 12 means 12号/小四.'),
    fontColor: z
      .string()
      .regex(/^#[0-9A-Fa-f]{6}$/)
      .optional()
      .describe('Font color as #RRGGBB.'),
    highlightColor: z.string().optional().describe('Highlight color as #RRGGBB.'),
    alignment: z
      .enum(['Left', 'Centered', 'Right', 'Justified'])
      .optional()
      .describe('Paragraph alignment (对齐方式). Justified = 两端对齐.'),
    lineSpacing: z
      .number()
      .positive()
      .max(1584)
      .optional()
      .describe('Exact line spacing in points (磅值). Prefer lineSpacingMultiple for 倍行距.'),
    lineSpacingMultiple: z
      .number()
      .positive()
      .max(10)
      .optional()
      .describe(
        'Relative line spacing (倍行距): 1 = 单倍, 1.5 = 1.5倍, 2 = 2倍. Converted to points using the font size.',
      ),
    spaceAfter: z.number().min(0).max(200).optional().describe('Spacing after paragraph in points (段后间距，磅).'),
  }

  if (shouldEnable('propose_format_patch')) {
    tools.push(
      tool(
        async input => {
          const args = input as Record<string, unknown> & { scope?: 'selection' | 'document' }
          const changes = restrictFormatChanges(extractFormatChanges(args), security?.allowedFormatFields)
          if (Object.keys(changes).length === 0) {
            throw new AppError('REQUEST_FAILED', 'At least one formatting change is required')
          }
          const scope = (args.scope || 'selection') as 'selection' | 'document'
          const anchor = scope === 'selection' ? await createSelectionAnchor() : null
          const beforeHash = anchor?.hash || (await readCurrentHash(scope))
          const request: FormatRequest = {
            id: crypto.randomUUID(),
            scope,
            changes,
            beforeHash,
            beforeOoxml: anchor?.ooxml || (scope === 'document' ? await readDocumentOoxml() : undefined),
            beforeOoxmlHash: anchor?.ooxml
              ? hashText(anchor.ooxml)
              : scope === 'document'
                ? hashText(await readDocumentOoxml())
                : undefined,
            anchorTag: anchor?.tag,
          }
          // One active proposal is exposed to the next agent turn. Creating a
          // replacement plan supersedes the previous unconfirmed plan.
          if (state.activeFormatRequestId) {
            const previous = state.formatRequests.get(state.activeFormatRequestId)
            if (previous?.anchorTag) void releaseSelectionAnchor(previous.anchorTag).catch(() => undefined)
            state.formatRequests.delete(state.activeFormatRequestId)
          }
          state.formatRequests.set(request.id, request)
          state.activeFormatRequestId = request.id
          return JSON.stringify({
            formatId: request.id,
            scope: request.scope,
            changes: request.changes,
            beforeHash: request.beforeHash,
            note: 'Present this formatting plan to the user and wait for their confirmation. After the user confirms (是/确定/执行/ok), call apply_format_patch with this formatId.',
          })
        },
        {
          name: 'propose_format_patch',
          description:
            'Create a reviewable formatting plan for the current selection or the whole document. Include only fields explicitly requested by the user and omit all unspecified fields. This does not alter visible content, but creates a hidden temporary anchor for selection scope. Present the plan, end the turn, then call apply_format_patch only if a later user message confirms it.',
          schema: z.object(formatFieldsSchema),
        },
      ),
    )
  }

  if (shouldEnable('apply_format_patch')) {
    tools.push(
      tool(
        async input => {
          const args = input as { formatId: string }
          const already = state.appliedFormatRequests.get(args.formatId)
          if (already) {
            const verification = await verifyFormatRequest(already)
            return JSON.stringify({
              formatId: args.formatId,
              status: 'already_applied',
              idempotent: true,
              verification,
            })
          }

          const request = state.formatRequests.get(args.formatId)
          if (!request || args.formatId !== state.activeFormatRequestId) {
            throw new AppError(
              'REQUEST_FAILED',
              'The format proposal was not found, is no longer active, or has expired',
            )
          }
          const currentHash = await readCurrentHash(request.scope)
          if (currentHash !== request.beforeHash) {
            state.formatRequests.delete(request.id)
            state.activeFormatRequestId = null
            if (request.anchorTag) void releaseSelectionAnchor(request.anchorTag).catch(() => undefined)
            throw new AppError('DOCUMENT_CONFLICT', 'The document changed while formatting was awaiting approval')
          }

          try {
            await applyFormatRequest(request)
          } catch (error) {
            if (request.anchorTag) void releaseSelectionAnchor(request.anchorTag).catch(() => undefined)
            throw error
          }
          request.appliedOoxmlHash = await readFormatOoxmlHash(request)
          state.appliedFormatRequests.set(request.id, request)
          state.formatRequests.delete(request.id)
          state.activeFormatRequestId = null
          security?.onFormatApplied?.(request)
          const verification = await verifyFormatRequest(request)
          return JSON.stringify({
            formatId: request.id,
            status: 'applied',
            scope: request.scope,
            changes: request.changes,
            verification,
          })
        },
        {
          name: 'apply_format_patch',
          description:
            'Apply the active formatting plan after you determine from the user message that they confirmed it. The operation is guarded by the active format ID and original content hash, then verifies the result. Do not call this in the proposal turn.',
          schema: z.object({ formatId: z.string().describe('The formatId returned by propose_format_patch') }),
        },
      ),
    )
  }

  if (shouldEnable('format_document_selection')) {
    tools.push(
      tool(
        async input => {
          const args = input as Record<string, unknown> & { scope?: 'selection' | 'document'; operationId?: string }
          const changes = restrictFormatChanges(extractFormatChanges(args), security?.allowedFormatFields)
          if (Object.keys(changes).length === 0) {
            throw new AppError('REQUEST_FAILED', 'At least one formatting change is required')
          }
          if (args.operationId && state.appliedFormatRequests.has(args.operationId)) {
            const applied = state.appliedFormatRequests.get(args.operationId)!
            return JSON.stringify({
              operationId: args.operationId,
              status: 'already_applied',
              idempotent: true,
              scope: applied.scope,
              changes: applied.changes,
              verification: await verifyFormatRequest(applied),
            })
          }
          const scope = (args.scope || 'selection') as 'selection' | 'document'
          const anchor = scope === 'selection' ? await createSelectionAnchor() : null
          const beforeHash = await readCurrentHash(scope)
          const request: FormatRequest = {
            id: (args.operationId as string) || crypto.randomUUID(),
            scope,
            changes,
            beforeHash,
            beforeOoxml: anchor?.ooxml || (scope === 'document' ? await readDocumentOoxml() : undefined),
            beforeOoxmlHash: anchor?.ooxml
              ? hashText(anchor.ooxml)
              : scope === 'document'
                ? hashText(await readDocumentOoxml())
                : undefined,
            anchorTag: anchor?.tag,
          }
          const currentHash = await readCurrentHash(request.scope)
          if (currentHash !== request.beforeHash) {
            if (request.anchorTag) void releaseSelectionAnchor(request.anchorTag).catch(() => undefined)
            throw new AppError('DOCUMENT_CONFLICT', 'The document changed while formatting was awaiting approval')
          }

          try {
            await applyFormatRequest(request)
          } catch (error) {
            if (request.anchorTag) void releaseSelectionAnchor(request.anchorTag).catch(() => undefined)
            throw error
          }
          request.appliedOoxmlHash = await readFormatOoxmlHash(request)
          state.appliedFormatRequests.set(request.id, request)
          if (state.activeFormatRequestId) {
            const previous = state.formatRequests.get(state.activeFormatRequestId)
            if (previous?.anchorTag) void releaseSelectionAnchor(previous.anchorTag).catch(() => undefined)
            state.formatRequests.delete(state.activeFormatRequestId)
          }
          state.activeFormatRequestId = null
          security?.onFormatApplied?.(request)
          const verification = await verifyFormatRequest(request)
          return JSON.stringify({
            operationId: request.id,
            status: 'applied',
            scope: request.scope,
            changes: request.changes,
            verification,
          })
        },
        {
          name: 'format_document_selection',
          description:
            'Immediately apply low-risk, reversible formatting to the current selection or whole document, then verify it. Include only fields explicitly requested by the user. Use propose_format_patch instead when the user explicitly asks to preview a plan or wait for confirmation.',
          schema: z.object({
            operationId: z
              .string()
              .optional()
              .describe('Reuse the operationId from a previous response to make a retry idempotent'),
            ...formatFieldsSchema,
          }),
        },
      ),
    )
  }

  if (shouldEnable('verify_patch')) {
    tools.push(
      tool(
        async input => {
          const args = input as { proposalId: string }
          const applied = state.appliedTextChanges.get(args.proposalId)
          if (!applied) {
            const proposal = state.textProposals.get(args.proposalId)
            return JSON.stringify({ proposalId: args.proposalId, status: proposal?.status || 'unknown' })
          }
          const verification = await verifyTextChange(applied)
          return JSON.stringify({ proposalId: args.proposalId, ...verification })
        },
        {
          name: 'verify_patch',
          description:
            'Verify that an applied patch anchor still exists and that the edited text was not changed afterward.',
          schema: z.object({ proposalId: z.string() }),
        },
      ),
    )
  }

  return tools
}

export const restoreAppliedTaskChange = async (change: AppliedTextChange) => restoreTextChange(change)

export const restoreFormatRequest = async (request: FormatRequest): Promise<void> => {
  const beforeOoxml = request.beforeOoxml
  if (!beforeOoxml) throw new AppError('REQUEST_FAILED', 'The original formatting snapshot is unavailable')
  await Word.run(async context => {
    const controls = request.anchorTag ? context.document.contentControls : undefined
    controls?.load('items/tag')
    if (controls) await context.sync()
    const range = request.anchorTag
      ? getAnchoredRange(controls!, request.anchorTag)
      : request.scope === 'document'
        ? context.document.body.getRange()
        : context.document.getSelection()
    if (!range) throw new AppError('DOCUMENT_CONFLICT', 'The formatting anchor is no longer available')
    const currentOoxml = range.getOoxml()
    range.load('text')
    await context.sync()
    if (request.beforeHash && hashText(range.text || '') !== request.beforeHash) {
      throw new AppError('DOCUMENT_CONFLICT', 'The document text changed after formatting was applied')
    }
    if (request.appliedOoxmlHash && hashText(currentOoxml.value || '') !== request.appliedOoxmlHash) {
      throw new AppError('DOCUMENT_CONFLICT', 'The document formatting changed after formatting was applied')
    }
    range.insertOoxml(beforeOoxml, 'Replace')
    await context.sync()
    if (request.anchorTag) {
      const control = controls!.items.find(item => item.tag === request.anchorTag)
      control?.delete(true)
      await context.sync()
    }
  })
}
