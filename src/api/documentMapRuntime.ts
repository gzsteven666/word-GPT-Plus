import { AppError } from '@/api/errors'
import {
  buildDocumentMap,
  DocumentMap,
  DocumentMapNode,
  getDocumentMapNode,
  isDocumentMapCurrent,
  selectDocumentMapNodes,
} from '@/utils/documentMap'
import { hashText } from '@/utils/textProposal'

export const canonicalizeDocumentOoxml = (ooxml: string): string =>
  ooxml
    .replace(/<\?xml[^>]*\?>/gi, '')
    .replace(/\s+(?:w|w\d+):(?:rsid[A-Za-z0-9]+|paraId|textId|docId|durableId)\s*=\s*(?:"[^"]*"|'[^']*')/gi, '')
    .replace(/>\s+</g, '><')
    .replace(/\s+/g, ' ')
    .trim()

export const hashDocumentOoxml = (ooxml: string): string => hashText(canonicalizeDocumentOoxml(ooxml))

const readBodySnapshot = async (context: Word.RequestContext) => {
  const body = context.document.body
  body.load('text')
  const ooxml = body.getRange().getOoxml()
  await context.sync()
  return {
    text: body.text || '',
    ooxml: ooxml.value || '',
  }
}

export const createDocumentMap = async (): Promise<DocumentMap> =>
  Word.run(async context => {
    const body = context.document.body
    const paragraphs = body.paragraphs
    paragraphs.load('items/text,styleBuiltIn,style')
    const snapshot = await readBodySnapshot(context)
    const paragraphInputs = paragraphs.items.map(paragraph => ({
      text: paragraph.text || '',
      styleBuiltIn: paragraph.styleBuiltIn,
      style: paragraph.style,
      protectedObjects: [],
      protectedObjectsAvailable: false,
    }))
    return buildDocumentMap({
      documentHash: hashText(snapshot.text),
      documentOoxmlHash: hashDocumentOoxml(snapshot.ooxml),
      paragraphs: paragraphInputs,
    })
  })

export const assertDocumentMapCurrent = async (context: Word.RequestContext, map: DocumentMap): Promise<void> => {
  const snapshot = await readBodySnapshot(context)
  if (!isDocumentMapCurrent(map, hashText(snapshot.text), hashDocumentOoxml(snapshot.ooxml)))
    throw new AppError('DOCUMENT_CONFLICT', 'The document map is stale; rebuild it before reading these nodes')
}

export interface DocumentMapNodeRead {
  nodeId: string
  kind: DocumentMapNode['kind']
  headingPath: string[]
  text: string
  textHash: string
  truncated: boolean
  protectedObjects: string[]
  protectedObjectsAvailable: boolean
}

export const readDocumentMapNodes = async (
  map: DocumentMap,
  nodeIds: string[],
  maxChars = 12000,
): Promise<DocumentMapNodeRead[]> =>
  Word.run(async context => {
    await assertDocumentMapCurrent(context, map)
    const paragraphs = context.document.body.paragraphs
    paragraphs.load('items/text,styleBuiltIn,style')
    await context.sync()
    const selected = selectDocumentMapNodes(map, nodeIds, maxChars)
    const results: DocumentMapNodeRead[] = []
    let remaining = maxChars
    for (const item of selected) {
      const node = getDocumentMapNode(map, item.node.id)
      if (!node) throw new AppError('DOCUMENT_CONFLICT', `Document map node is no longer available: ${item.node.id}`)
      const paragraph = paragraphs.items[node.order]
      if (!paragraph) throw new AppError('DOCUMENT_CONFLICT', `Document map node moved: ${item.node.id}`)
      const text = paragraph.text || ''
      if (hashText(text) !== node.textHash)
        throw new AppError('DOCUMENT_CONFLICT', `Document map node changed: ${item.node.id}`)
      const value = text.slice(0, remaining)
      results.push({
        nodeId: node.id,
        kind: node.kind,
        headingPath: node.headingPath,
        text: value,
        textHash: hashText(text),
        truncated: value.length < text.length,
        protectedObjects: node.protectedObjects,
        protectedObjectsAvailable: node.protectedObjectsAvailable,
      })
      remaining -= value.length
      if (remaining <= 0) break
    }
    return results
  })

export const assertDocumentMapIsCurrent = async (map: DocumentMap): Promise<void> =>
  Word.run(async context => {
    await assertDocumentMapCurrent(context, map)
  })
