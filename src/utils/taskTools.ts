import { tool } from '@langchain/core/tools'
import { z } from 'zod'

import { AppError } from '@/api/errors'
import {
  AppliedTextChange,
  applyTextChangeProposal,
  makeSelectionProposal,
  readSelectionSnapshot,
  restoreTextChange,
  verifyTextChange,
} from '@/api/safeEdit'
import { buildTextDiff, createTextChangeProposal, hashText, TextChangeProposal } from '@/utils/textProposal'

export type TaskToolName =
  | 'read_document_structure'
  | 'read_range'
  | 'propose_document_patch'
  | 'apply_document_patch'
  | 'propose_format_patch'
  | 'apply_format_patch'
  | 'format_document_selection'
  | 'verify_patch'

export interface FormatRequest {
  id: string
  scope: 'selection' | 'document'
  changes: Record<string, boolean | number | string>
  beforeHash?: string
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
  appliedFormatRequests: Map<string, FormatRequest>
}

export const createTaskToolState = (): TaskToolState => ({
  textProposals: new Map(),
  appliedTextChanges: new Map(),
  formatRequests: new Map(),
  appliedFormatRequests: new Map(),
})

export interface TaskToolSecurity {
  requestTextChangeApproval?: (proposal: TextChangeProposal) => Promise<boolean>
  requestFormatApproval?: (request: FormatRequest) => Promise<boolean>
  requestSensitiveDataApproval?: (scope: string) => Promise<boolean>
  onTextChangeApplied?: (change: AppliedTextChange) => void
}

/**
 * Detect a short confirmation/continuation reply (是/好/确定/执行/yes/ok/...).
 * Such replies carry no task keywords of their own, so the agent needs the
 * write tools from the previous turn to stay available.
 */
export const isConfirmationIntent = (text: string): boolean =>
  /^(是|好|好的|确定|确认|可以|执行|应用|继续|同意|没问题|选中了|已选中|完成|yes|ok|okay|confirm|apply|proceed|go ahead|sure|selected|done|accept|fine)[！!。.，,\s]*$/i.test(
    text.trim(),
  )

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
  if (operation === 'replace' || operation === 'delete') {
    return makeSelectionProposal(afterText, 'agent', operation)
  }

  const snapshot = await readSelectionSnapshot()
  return createTextChangeProposal({
    operation,
    beforeText: snapshot.text,
    afterText: `${snapshot.text}${afterText}`,
    beforeOoxml: snapshot.ooxml,
    source: 'agent',
  })
}

const readDocumentHash = async (): Promise<string> =>
  Word.run(async context => {
    const body = context.document.body
    body.load('text')
    await context.sync()
    return hashText(body.text || '')
  })

const readCurrentHash = async (scope: FormatRequest['scope']): Promise<string> =>
  scope === 'selection' ? (await readSelectionSnapshot()).hash : readDocumentHash()

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

const applyFormatRequest = async (request: FormatRequest): Promise<void> => {
  const changes = request.changes
  await Word.run(async context => {
    const range = request.scope === 'document' ? context.document.body.getRange() : context.document.getSelection()

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
    const range = request.scope === 'document' ? context.document.body.getRange() : context.document.getSelection()
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
          const snapshot = await readSelectionSnapshot()
          return JSON.stringify({
            text: snapshot.text,
            hash: snapshot.hash,
            protectedObjects: snapshot.protectedObjects,
          })
        },
        {
          name: 'read_range',
          description:
            'Read the currently selected Word range, its content hash, and protected-object markers. This is read-only.',
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
            'Create a reviewable text patch for the current Word selection. This never writes to the document. Call apply_document_patch only after the proposal is ready for user review.',
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
          const changes = extractFormatChanges(args)
          if (Object.keys(changes).length === 0) {
            throw new AppError('REQUEST_FAILED', 'At least one formatting change is required')
          }
          const scope = (args.scope || 'selection') as 'selection' | 'document'
          const beforeHash = await readCurrentHash(scope)
          const request: FormatRequest = {
            id: crypto.randomUUID(),
            scope,
            changes,
            beforeHash,
          }
          state.formatRequests.set(request.id, request)
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
            'Create a reviewable formatting plan for the current selection or the whole document. This never writes to the document and never shows a dialog. Present the plan to the user, then call apply_format_patch after the user confirms.',
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
          if (!request) throw new AppError('REQUEST_FAILED', 'The format proposal was not found or has expired')
          if (!security?.requestFormatApproval) {
            throw new AppError('REQUEST_FAILED', 'A user approval handler is required before applying formatting')
          }
          if (!(await security.requestFormatApproval(request))) {
            return JSON.stringify({ formatId: request.id, status: 'cancelled' })
          }

          const currentHash = await readCurrentHash(request.scope)
          if (currentHash !== request.beforeHash) {
            throw new AppError('DOCUMENT_CONFLICT', 'The document changed while formatting was awaiting approval')
          }

          await applyFormatRequest(request)
          state.appliedFormatRequests.set(request.id, request)
          state.formatRequests.delete(request.id)
          const verification = await verifyFormatRequest(request)
          return JSON.stringify({ formatId: request.id, status: 'applied', verification })
        },
        {
          name: 'apply_format_patch',
          description:
            'Apply a previously proposed formatting plan only after explicit user approval. The operation is guarded by the original content hash and format ID, then verifies the result was applied.',
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
          const changes = extractFormatChanges(args)
          if (Object.keys(changes).length === 0) {
            throw new AppError('REQUEST_FAILED', 'At least one formatting change is required')
          }
          if (args.operationId && state.appliedFormatRequests.has(args.operationId)) {
            return JSON.stringify({ operationId: args.operationId, status: 'already_applied', idempotent: true })
          }
          const scope = (args.scope || 'selection') as 'selection' | 'document'
          const beforeHash = await readCurrentHash(scope)
          const request: FormatRequest = {
            id: (args.operationId as string) || crypto.randomUUID(),
            scope,
            changes,
            beforeHash,
          }
          if (!security?.requestFormatApproval) {
            throw new AppError('REQUEST_FAILED', 'A user approval handler is required before applying formatting')
          }
          if (!(await security.requestFormatApproval(request))) {
            return JSON.stringify({ operationId: request.id, status: 'cancelled' })
          }
          const currentHash = await readCurrentHash(request.scope)
          if (currentHash !== request.beforeHash) {
            throw new AppError('DOCUMENT_CONFLICT', 'The document changed while formatting was awaiting approval')
          }

          await applyFormatRequest(request)
          state.appliedFormatRequests.set(request.id, request)
          const verification = await verifyFormatRequest(request)
          return JSON.stringify({ operationId: request.id, status: 'applied', verification })
        },
        {
          name: 'format_document_selection',
          description:
            'Apply explicit, user-approved formatting to the current selection or the whole document in a single step (shows an approval dialog). This changes styles only and never changes text content. Prefer propose_format_patch + apply_format_patch when the user wants to review the plan first.',
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
