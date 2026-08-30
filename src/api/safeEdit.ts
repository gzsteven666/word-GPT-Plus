import { AppError } from '@/api/errors'
import { createTextChangeProposal, hashText, TextChangeProposal } from '@/utils/textProposal'

export interface SelectionSnapshot {
  text: string
  ooxml: string
  hash: string
  protectedObjects: string[]
}

export interface AppliedTextChange {
  proposal: TextChangeProposal
  bookmarkTag: string
  appliedHash: string
}

const protectedObjectNames = (ooxml: string): string[] => {
  const rules: [string, RegExp][] = [
    ['图片或绘图', /<(?:wp:)?anchor|<pic:pic|<w:drawing/i],
    ['表格', /<w:tbl\b/i],
    ['公式', /<m:oMath|<m:oMathPara/i],
    ['内容控件', /<w:sdt\b/i],
    ['脚注或尾注', /<w:(?:footnoteReference|endnoteReference)\b/i],
    ['域或对象', /<w:(?:object|fldSimple|instrText)\b/i],
  ]
  return rules.filter(([, rule]) => rule.test(ooxml)).map(([name]) => name)
}

export const readSelectionSnapshot = async (): Promise<SelectionSnapshot> =>
  Word.run(async context => {
    const range = context.document.getSelection()
    const ooxml = range.getOoxml()
    range.load('text')
    await context.sync()
    const ooxmlText = ooxml.value || ''
    return {
      text: range.text || '',
      ooxml: ooxmlText,
      hash: hashText(range.text || ''),
      protectedObjects: protectedObjectNames(ooxmlText),
    }
  })

export const makeSelectionProposal = async (
  afterText: string,
  source: TextChangeProposal['source'],
  operation: TextChangeProposal['operation'] = 'replace',
): Promise<TextChangeProposal> => {
  const snapshot = await readSelectionSnapshot()
  if (!snapshot.text && operation === 'replace') {
    throw new AppError('WORD_API_UNSUPPORTED', 'No text is selected')
  }
  if (snapshot.protectedObjects.length > 0) {
    throw new AppError(
      'NON_TEXT_OBJECTS',
      `The selection contains protected objects: ${snapshot.protectedObjects.join(', ')}`,
      { details: { objectCount: snapshot.protectedObjects.length } },
    )
  }
  return createTextChangeProposal({
    operation,
    beforeText: snapshot.text,
    afterText,
    beforeOoxml: snapshot.ooxml,
    source,
  })
}

export const applyTextChangeProposal = async (proposal: TextChangeProposal): Promise<AppliedTextChange> => {
  if (proposal.status !== 'pending') {
    throw new AppError('REQUEST_FAILED', 'This edit proposal is no longer pending')
  }

  return Word.run(async context => {
    const bookmarkTag = `wordgpt_${proposal.id.replaceAll('-', '')}`
    const existingControls = context.document.contentControls
    existingControls.load('items/tag')
    await context.sync()
    const existing = existingControls.items.find(item => item.tag === bookmarkTag)
    if (existing) {
      const existingRange = existing.getRange('Content')
      existingRange.load('text')
      await context.sync()
      proposal.status = 'applied'
      return {
        proposal,
        bookmarkTag,
        appliedHash: hashText(existingRange.text || proposal.afterText),
      }
    }

    const range = context.document.getSelection()
    const ooxml = range.getOoxml()
    range.load('text')
    await context.sync()

    const currentText = range.text || ''
    if (hashText(currentText) !== proposal.expectedBeforeHash) {
      throw new AppError('DOCUMENT_CONFLICT', 'The selection changed before the edit was applied')
    }

    const currentObjects = protectedObjectNames(ooxml.value || '')
    if (currentObjects.length > 0) {
      throw new AppError('NON_TEXT_OBJECTS', `The selection contains protected objects: ${currentObjects.join(', ')}`)
    }

    const control = range.insertContentControl()
    control.tag = bookmarkTag
    control.title = 'Word GPT edit'
    control.appearance = 'Hidden'
    const contentRange = control.getRange('Content')
    contentRange.insertText(proposal.afterText, 'Replace')
    contentRange.load('text')
    await context.sync()

    proposal.status = 'applied'
    return {
      proposal,
      bookmarkTag,
      appliedHash: hashText(contentRange.text || proposal.afterText),
    }
  })
}

export const verifyTextChange = async (
  change: AppliedTextChange,
): Promise<{ status: 'verified' | 'changed'; currentHash?: string }> =>
  Word.run(async context => {
    const controls = context.document.contentControls
    controls.load('items/tag')
    await context.sync()

    const control = controls.items.find(item => item.tag === change.bookmarkTag)
    if (!control) throw new AppError('DOCUMENT_CONFLICT', 'The patch anchor is no longer available')

    const contentRange = control.getRange('Content')
    contentRange.load('text')
    await context.sync()
    const currentHash = hashText(contentRange.text || '')
    return {
      status: currentHash === change.appliedHash ? 'verified' : 'changed',
      currentHash,
    }
  })

export const restoreTextChange = async (change: AppliedTextChange): Promise<TextChangeProposal> =>
  Word.run(async context => {
    const controls = context.document.contentControls
    controls.load('items')
    await context.sync()

    const control = controls.items.find(item => item.tag === change.bookmarkTag)
    if (!control) throw new AppError('DOCUMENT_CONFLICT', 'The previous edit anchor is no longer available')

    const contentRange = control.getRange('Content')
    contentRange.load('text')
    await context.sync()
    if (hashText(contentRange.text || '') !== change.appliedHash) {
      throw new AppError('DOCUMENT_CONFLICT', 'The edited text changed after it was applied')
    }

    contentRange.insertText(change.proposal.beforeText, 'Replace')
    contentRange.load('text')
    await context.sync()
    change.proposal.status = 'restored'
    return change.proposal
  })
