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
import { buildTextDiff, createTextChangeProposal, TextChangeProposal } from '@/utils/textProposal'

export type TaskToolName =
  | 'read_document_structure'
  | 'read_range'
  | 'propose_document_patch'
  | 'apply_document_patch'
  | 'verify_patch'

export interface TaskToolSecurity {
  requestTextChangeApproval?: (proposal: TextChangeProposal) => Promise<boolean>
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
