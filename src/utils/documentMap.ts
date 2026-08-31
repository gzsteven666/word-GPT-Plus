import { hashText } from '@/utils/textProposal'

export type DocumentMapNodeKind = 'heading' | 'paragraph'

export interface DocumentMapParagraphInput {
  text: string
  styleBuiltIn?: string
  style?: string
  ooxml?: string
  protectedObjects?: string[]
  protectedObjectsAvailable?: boolean
}

export interface DocumentMapNode {
  id: string
  kind: DocumentMapNodeKind
  order: number
  headingLevel?: number
  headingPath: string[]
  preview: string
  textHash: string
  ooxmlHash?: string
  beforeNodeHash?: string
  afterNodeHash?: string
  protectedObjects: string[]
  protectedObjectsAvailable: boolean
}

export interface DocumentMap {
  id: string
  documentHash: string
  documentOoxmlHash: string
  nodes: DocumentMapNode[]
  createdAt: string
}

export interface DocumentMapSearchResult {
  node: DocumentMapNode
  score: number
}

export interface DocumentMapNodeSelection {
  node: DocumentMapNode
  text: string
}

const PREVIEW_LIMIT = 160

export const isDocumentMapIntent = (text: string): boolean =>
  /文档地图|建图|标题树|章节检索|长文档|document\s*map|heading\s*tree|long\s*document|map\s*id|target\s*node\s*id/i.test(
    text,
  )

export const hasDocumentMapTargetReference = (text: string): boolean =>
  /map\s*id\s*[:：`\s]*map-[a-f0-9]+/i.test(text) && /(?:target\s*)?node\s*id\s*[:：`\s]*node-/i.test(text)

const headingLevelFrom = (paragraph: DocumentMapParagraphInput): number | undefined => {
  const style = paragraph.styleBuiltIn || paragraph.style || ''
  const match = style.match(/^(?:Heading|标题)\s*([1-9])$/i)
  return match ? Number(match[1]) : undefined
}

const previewOf = (text: string): string => text.replace(/\s+/g, ' ').trim().slice(0, PREVIEW_LIMIT)

const nodeFingerprint = (kind: DocumentMapNodeKind, order: number, text: string, headingPath: string[]): string =>
  hashText(`${kind}|${order}|${headingPath.join(' / ')}|${text}`).slice(0, 12)

export const buildDocumentMap = (input: {
  documentHash: string
  documentOoxmlHash: string
  paragraphs: DocumentMapParagraphInput[]
}): DocumentMap => {
  const headingPath: string[] = []
  const nodes: DocumentMapNode[] = input.paragraphs.map((paragraph, order) => {
    const text = paragraph.text || ''
    const headingLevel = headingLevelFrom(paragraph)
    if (headingLevel) {
      headingPath.splice(headingLevel - 1)
      headingPath[headingLevel - 1] = text.trim()
      headingPath.length = headingLevel
    }
    const kind: DocumentMapNodeKind = headingLevel ? 'heading' : 'paragraph'
    const path = headingPath.filter(Boolean)
    return {
      id: `node-${order + 1}-${nodeFingerprint(kind, order, text, path)}`,
      kind,
      order,
      ...(headingLevel ? { headingLevel } : {}),
      headingPath: [...path],
      preview: previewOf(text),
      textHash: hashText(text),
      ...(paragraph.ooxml ? { ooxmlHash: hashText(paragraph.ooxml) } : {}),
      protectedObjects: [...(paragraph.protectedObjects || [])],
      protectedObjectsAvailable: paragraph.protectedObjectsAvailable ?? Boolean(paragraph.ooxml),
    } satisfies DocumentMapNode
  })

  nodes.forEach((node, index) => {
    node.beforeNodeHash = nodes[index - 1]?.textHash
    node.afterNodeHash = nodes[index + 1]?.textHash
  })

  return {
    id: `map-${hashText(`${input.documentHash}|${input.documentOoxmlHash}|${nodes.length}`).slice(0, 16)}`,
    documentHash: input.documentHash,
    documentOoxmlHash: input.documentOoxmlHash,
    nodes,
    createdAt: new Date().toISOString(),
  }
}

const queryTerms = (query: string): string[] =>
  query
    .toLocaleLowerCase()
    .trim()
    .split(/[\s,，。；;、]+/u)
    .filter(Boolean)

export const queryDocumentMap = (map: DocumentMap, query: string, limit = 20): DocumentMapSearchResult[] => {
  const normalized = query.toLocaleLowerCase().trim()
  if (!normalized || limit <= 0) return []
  const terms = queryTerms(normalized)
  return map.nodes
    .map(node => {
      const haystack = `${node.headingPath.join(' / ')}\n${node.preview}`.toLocaleLowerCase()
      const matchedTerms = terms.filter(term => haystack.includes(term)).length
      const exactBonus = normalized && haystack.includes(normalized) ? terms.length : 0
      return { node, score: matchedTerms + exactBonus }
    })
    .filter(result => result.score > 0)
    .sort((left, right) => right.score - left.score || left.node.order - right.node.order)
    .slice(0, limit)
}

export const selectDocumentMapNodes = (
  map: DocumentMap,
  nodeIds: string[],
  maxChars = 12000,
): DocumentMapNodeSelection[] => {
  if (maxChars <= 0) return []
  const selected: DocumentMapNodeSelection[] = []
  let remaining = maxChars
  for (const id of nodeIds) {
    const node = map.nodes.find(item => item.id === id)
    if (!node || remaining <= 0) continue
    const text = node.preview.slice(0, remaining)
    if (!text) continue
    selected.push({ node, text })
    remaining -= text.length
  }
  return selected
}

export const getDocumentMapNode = (map: DocumentMap, nodeId: string): DocumentMapNode | undefined =>
  map.nodes.find(node => node.id === nodeId)

export const isDocumentMapCurrent = (map: DocumentMap, documentHash: string, documentOoxmlHash: string): boolean =>
  map.documentHash === documentHash && map.documentOoxmlHash === documentOoxmlHash
