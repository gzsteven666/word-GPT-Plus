import assert from 'node:assert/strict'

import { proposeDocumentPatchSet } from '../src/api/documentPatchRuntime.ts'
import { buildDocumentMap } from '../src/utils/documentMap.ts'
import { hashText } from '../src/utils/textProposal.ts'

const originalWord = globalThis.Word

try {
  const paragraphTexts = ['第一章', '目标文本', '第二章', '目标文本']
  const bodyText = paragraphTexts.join('\n')
  const bodyOoxml = `<w:body><w:p><w:r><w:t>${bodyText}</w:t></w:r></w:p></w:body>`
  const makeTargetRange = (text: string) => {
    const range: any = {
      text,
      ooxml: `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`,
      load: () => undefined,
      getOoxml: () => ({ value: range.ooxml }),
      getRange: () => range,
      search: (query: string) => ({ items: query === text ? [range] : [], load: () => undefined }),
      compareLocationWith: () => ({ value: 'Before' }),
      insertContentControl: () => ({
        tag: '',
        title: '',
        appearance: '',
        getRange: () => range,
        delete: () => undefined,
      }),
    }
    return range
  }
  const paragraphRanges = paragraphTexts.map(text => {
    const range = makeTargetRange(text)
    return {
      text,
      styleBuiltIn: text.endsWith('章') ? 'Heading1' : 'Normal',
      load: () => undefined,
      getRange: () => range,
      search: (query: string) => ({
        items: query === text ? [range] : [],
        load: () => undefined,
      }),
    }
  })
  const body = {
    text: bodyText,
    load: () => undefined,
    getRange: () => ({ getOoxml: () => ({ value: bodyOoxml }) }),
    paragraphs: {
      items: paragraphRanges,
      load: () => undefined,
    },
    search: () => ({ items: [], load: () => undefined }),
  }
  const context = { document: { body }, sync: async () => undefined }
  globalThis.Word = {
    run: async callback => callback(context as unknown as Word.RequestContext),
  } as unknown as typeof Word

  const map = buildDocumentMap({
    documentHash: hashText(bodyText),
    documentOoxmlHash: hashText(bodyOoxml.replace(/\s+/g, ' ').trim()),
    paragraphs: paragraphTexts.map(text => ({ text, styleBuiltIn: text.endsWith('章') ? 'Heading1' : 'Normal' })),
  })
  const targetNode = map.nodes[1]
  const patchSet = await proposeDocumentPatchSet(
    [
      {
        type: 'replace',
        targetText: '目标文本',
        replacementText: '已修改',
        mapId: map.id,
        targetNodeId: targetNode.id,
      },
    ],
    map,
  )
  assert.equal(patchSet.operations[0].beforeText, '目标文本')

  await assert.rejects(
    () =>
      proposeDocumentPatchSet(
        [
          {
            type: 'replace',
            targetText: '目标文本',
            replacementText: '错误',
            mapId: map.id,
            targetNodeId: 'missing-node',
          },
        ],
        map,
      ),
    /not available/i,
  )
} finally {
  globalThis.Word = originalWord
}

console.log('document patch node runtime tests: PASS')
