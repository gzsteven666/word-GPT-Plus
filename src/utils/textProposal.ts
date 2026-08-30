import { AppError } from '@/api/errors'

export type TextChangeOperation = 'replace' | 'insert' | 'delete'
export type TextChangeRisk = 'low' | 'medium' | 'high'
export type TextChangeSource = 'quick-action' | 'chat' | 'agent'
export type TextChangeStatus = 'pending' | 'applied' | 'cancelled' | 'restored'

export interface TextChangeProposal {
  id: string
  operation: TextChangeOperation
  scope: 'selection'
  beforeText: string
  afterText: string
  beforeOoxml?: string
  expectedBeforeHash: string
  risk: TextChangeRisk
  source: TextChangeSource
  status: TextChangeStatus
  createdAt: string
  protectedObjects?: string[]
}

export interface DiffLine {
  type: 'added' | 'removed' | 'unchanged'
  text: string
}

const tokenize = (text: string): string[] => {
  if (typeof Intl !== 'undefined' && 'Segmenter' in Intl) {
    const segmenter = new Intl.Segmenter(undefined, { granularity: 'word' })
    return [...segmenter.segment(text)].map(item => item.segment)
  }
  return text.match(/\s+|[^\s]+/g) || []
}

export const hashText = (text: string): string => {
  let hash = 2166136261
  for (let index = 0; index < text.length; index += 1) {
    hash ^= text.charCodeAt(index)
    hash = Math.imul(hash, 16777619)
  }
  return (hash >>> 0).toString(16).padStart(8, '0')
}

export const getProtectedObjects = (ooxml: string): string[] => {
  const objectRules: [string, RegExp][] = [
    ['图片或绘图', /<(?:wp:)?anchor|<pic:pic|<w:drawing/i],
    ['表格', /<w:tbl\b/i],
    ['公式', /<m:oMath|<m:oMathPara/i],
    ['内容控件', /<w:sdt\b/i],
    ['脚注或尾注', /<w:(?:footnoteReference|endnoteReference)\b/i],
    ['域或对象', /<w:(?:object|fldSimple|instrText)\b/i],
  ]
  return objectRules.filter(([, rule]) => rule.test(ooxml)).map(([label]) => label)
}

const calculateRisk = (beforeText: string, afterText: string): TextChangeRisk => {
  const beforeLength = beforeText.length
  const changedRatio = beforeLength === 0 ? 1 : Math.abs(afterText.length - beforeLength) / beforeLength
  if (beforeLength > 2000 || changedRatio > 0.75) return 'high'
  if (beforeLength > 500 || changedRatio > 0.35) return 'medium'
  return 'low'
}

export const createTextChangeProposal = (options: {
  operation: TextChangeOperation
  beforeText: string
  afterText: string
  beforeOoxml?: string
  source: TextChangeSource
}): TextChangeProposal => {
  if (options.beforeOoxml) {
    const protectedObjects = getProtectedObjects(options.beforeOoxml)
    if (protectedObjects.length > 0) {
      throw new AppError(
        'NON_TEXT_OBJECTS',
        `The selection contains protected objects: ${protectedObjects.join(', ')}`,
        { details: { objectCount: protectedObjects.length } },
      )
    }
  }

  return {
    id: crypto.randomUUID(),
    operation: options.operation,
    scope: 'selection',
    beforeText: options.beforeText,
    afterText: options.afterText,
    beforeOoxml: options.beforeOoxml,
    expectedBeforeHash: hashText(options.beforeText),
    risk: calculateRisk(options.beforeText, options.afterText),
    source: options.source,
    status: 'pending',
    createdAt: new Date().toISOString(),
  }
}

export const buildTextDiff = (beforeText: string, afterText: string): DiffLine[] => {
  const before = tokenize(beforeText)
  const after = tokenize(afterText)

  if (before.length * after.length > 250000) {
    return beforeText === afterText
      ? [{ type: 'unchanged', text: beforeText }]
      : [
          ...(beforeText ? [{ type: 'removed' as const, text: beforeText }] : []),
          ...(afterText ? [{ type: 'added' as const, text: afterText }] : []),
        ]
  }

  const table: number[][] = Array.from({ length: before.length + 1 }, () => Array(after.length + 1).fill(0))
  for (let row = before.length - 1; row >= 0; row -= 1) {
    for (let column = after.length - 1; column >= 0; column -= 1) {
      table[row][column] =
        before[row] === after[column]
          ? table[row + 1][column + 1] + 1
          : Math.max(table[row + 1][column], table[row][column + 1])
    }
  }

  const lines: DiffLine[] = []
  let row = 0
  let column = 0
  const append = (type: DiffLine['type'], text: string) => {
    if (!text) return
    const last = lines.at(-1)
    if (last?.type === type) last.text += text
    else lines.push({ type, text })
  }

  while (row < before.length && column < after.length) {
    if (before[row] === after[column]) {
      append('unchanged', before[row])
      row += 1
      column += 1
    } else if (table[row + 1][column] >= table[row][column + 1]) {
      append('removed', before[row])
      row += 1
    } else {
      append('added', after[column])
      column += 1
    }
  }
  while (row < before.length) append('removed', before[row++])
  while (column < after.length) append('added', after[column++])
  return lines
}

export const countDiffChanges = (diff: DiffLine[]) => ({
  added: diff.filter(line => line.type === 'added').length,
  removed: diff.filter(line => line.type === 'removed').length,
})
