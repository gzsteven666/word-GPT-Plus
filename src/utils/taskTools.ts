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
  | 'format_document_selection'
  | 'verify_patch'

export interface FormatRequest {
  id: string
  scope: 'selection' | 'document'
  changes: Record<string, boolean | number | string>
  beforeHash?: string
}

export interface TaskToolSecurity {
  requestTextChangeApproval?: (proposal: TextChangeProposal) => Promise<boolean>
  requestFormatApproval?: (request: FormatRequest) => Promise<boolean>
  requestSensitiveDataApproval?: (scope: string) => Promise<boolean>
  onTextChangeApplied?: (change: AppliedTextChange) => void
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

export const createTaskTools = (enabledTools?: TaskToolName[], security?: TaskToolSecurity) => {
  const proposals = new Map<string, TextChangeProposal>()
  const appliedChanges = new Map<string, AppliedTextChange>()
  const appliedFormatRequests = new Map<string, FormatRequest>()
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
          proposals.set(proposal.id, proposal)
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
          const existing = appliedChanges.get(args.proposalId)
          if (existing) {
            return JSON.stringify({ proposalId: args.proposalId, status: 'already_applied', idempotent: true })
          }

          const proposal = proposals.get(args.proposalId)
          if (!proposal) throw new AppError('REQUEST_FAILED', 'The patch proposal was not found or has expired')
          if (!security?.requestTextChangeApproval) {
            throw new AppError('REQUEST_FAILED', 'A user approval handler is required before applying a document patch')
          }
          if (!(await security.requestTextChangeApproval(proposal))) {
            return JSON.stringify({ proposalId: proposal.id, status: 'cancelled' })
          }

          const applied = await applyTextChangeProposal(proposal)
          appliedChanges.set(proposal.id, applied)
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

  if (shouldEnable('format_document_selection')) {
    tools.push(
      tool(
        async input => {
          const args = input as {
            scope: 'selection' | 'document'
            operationId?: string
            bold?: boolean
            italic?: boolean
            underline?: boolean
            fontName?: string
            fontSize?: number
            fontColor?: string
            highlightColor?: string
            alignment?: 'Left' | 'Centered' | 'Right' | 'Justified'
            lineSpacing?: number
            spaceAfter?: number
          }
          const changes = Object.fromEntries(
            Object.entries(args).filter(
              ([key, value]) => !['scope', 'operationId'].includes(key) && value !== undefined,
            ),
          ) as Record<string, boolean | number | string>
          if (Object.keys(changes).length === 0) {
            throw new AppError('REQUEST_FAILED', 'At least one formatting change is required')
          }
          if (args.operationId && appliedFormatRequests.has(args.operationId)) {
            return JSON.stringify({ operationId: args.operationId, status: 'already_applied', idempotent: true })
          }
          const beforeHash =
            args.scope === 'selection'
              ? (await readSelectionSnapshot()).hash
              : await Word.run(async context => {
                  const body = context.document.body
                  body.load('text')
                  await context.sync()
                  return hashText(body.text || '')
                })
          const request: FormatRequest = {
            id: args.operationId || crypto.randomUUID(),
            scope: args.scope,
            changes,
            beforeHash,
          }
          if (!security?.requestFormatApproval) {
            throw new AppError('REQUEST_FAILED', 'A user approval handler is required before applying formatting')
          }
          if (!(await security.requestFormatApproval(request))) {
            return JSON.stringify({ operationId: request.id, status: 'cancelled' })
          }
          const currentHash =
            request.scope === 'selection'
              ? (await readSelectionSnapshot()).hash
              : await Word.run(async context => {
                  const body = context.document.body
                  body.load('text')
                  await context.sync()
                  return hashText(body.text || '')
                })
          if (currentHash !== request.beforeHash) {
            throw new AppError('DOCUMENT_CONFLICT', 'The document changed while formatting was awaiting approval')
          }

          await Word.run(async context => {
            const range =
              request.scope === 'document' ? context.document.body.getRange() : context.document.getSelection()
            if (args.bold !== undefined) range.font.bold = args.bold
            if (args.italic !== undefined) range.font.italic = args.italic
            if (args.underline !== undefined) range.font.underline = args.underline ? 'Single' : 'None'
            if (args.fontName) range.font.name = args.fontName
            if (args.fontSize !== undefined) range.font.size = args.fontSize
            if (args.fontColor) range.font.color = args.fontColor
            if (args.highlightColor) range.font.highlightColor = args.highlightColor
            if (args.alignment || args.lineSpacing !== undefined || args.spaceAfter !== undefined) {
              const paragraphs = range.paragraphs
              paragraphs.load('items')
              await context.sync()
              for (const paragraph of paragraphs.items) {
                if (args.alignment) paragraph.alignment = args.alignment as Word.Alignment
                if (args.lineSpacing !== undefined) paragraph.lineSpacing = args.lineSpacing
                if (args.spaceAfter !== undefined) paragraph.spaceAfter = args.spaceAfter
              }
            }
            await context.sync()
          })
          appliedFormatRequests.set(request.id, request)
          return JSON.stringify({ operationId: request.id, status: 'applied', changes: request.changes })
        },
        {
          name: 'format_document_selection',
          description:
            'Apply explicit, user-approved formatting to the current selection or the whole document. This changes styles only and never changes text content.',
          schema: z.object({
            scope: z.enum(['selection', 'document']).default('selection'),
            operationId: z
              .string()
              .optional()
              .describe('Reuse the operationId from a previous response to make a retry idempotent'),
            bold: z.boolean().optional(),
            italic: z.boolean().optional(),
            underline: z.boolean().optional(),
            fontName: z.string().optional(),
            fontSize: z.number().positive().max(200).optional(),
            fontColor: z
              .string()
              .regex(/^#[0-9A-Fa-f]{6}$/)
              .optional(),
            highlightColor: z.string().optional(),
            alignment: z.enum(['Left', 'Centered', 'Right', 'Justified']).optional(),
            lineSpacing: z.number().positive().max(10).optional(),
            spaceAfter: z.number().min(0).max(200).optional(),
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
          const applied = appliedChanges.get(args.proposalId)
          if (!applied) {
            const proposal = proposals.get(args.proposalId)
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
