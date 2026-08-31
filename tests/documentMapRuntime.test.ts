import assert from 'node:assert/strict'

import {
  canonicalizeDocumentOoxml,
  createDocumentMap,
  hashDocumentOoxml,
  readDocumentMapNodes,
} from '../src/api/documentMapRuntime.ts'

const originalWord = globalThis.Word

try {
  const paragraphs = [
    { text: '项目报告', styleBuiltIn: 'Heading1' },
    { text: '收入同比增长 20%。', styleBuiltIn: 'Normal' },
    { text: '风险与计划', styleBuiltIn: 'Heading2' },
    { text: '下一步完成验收。', styleBuiltIn: 'Normal' },
  ]
  let bodyText = paragraphs.map(item => item.text).join('\n')
  const bodyOoxml = () => `<w:body><w:p><w:r><w:t>${bodyText}</w:t></w:r></w:p></w:body>`

  const context = {
    document: {
      body: {
        text: bodyText,
        load: () => undefined,
        getRange: () => ({ getOoxml: () => ({ value: bodyOoxml() }) }),
        paragraphs: {
          items: paragraphs.map(item => ({ ...item, load: () => undefined })),
          load: () => undefined,
        },
      },
    },
    sync: async () => undefined,
  }
  globalThis.Word = {
    run: async callback => callback(context as unknown as Word.RequestContext),
  } as unknown as typeof Word

  const map = await createDocumentMap()
  assert.equal(map.nodes.length, 4)
  assert.deepEqual(map.nodes[2].headingPath, ['项目报告', '风险与计划'])

  const read = await readDocumentMapNodes(map, [map.nodes[1].id, map.nodes[3].id], 20)
  assert.deepEqual(
    read.map(item => item.text),
    ['收入同比增长 20%。', '下一步完成验收。'],
  )

  bodyText = `${bodyText}\n新增段落`
  await assert.rejects(() => readDocumentMapNodes(map, [map.nodes[1].id]), /changed|stale|conflict/i)

  const stableOoxml = '<w:body><w:p w:rsidR="aaa" w15:paraId="111"><w:r><w:t>文本</w:t></w:r></w:p></w:body>'
  const volatileOoxml =
    '<?xml version="1.0"?><w:body w:rsidR="bbb" w15:paraId="222"><w:p><w:r><w:t>文本</w:t></w:r></w:p></w:body>'
  assert.equal(canonicalizeDocumentOoxml(stableOoxml), canonicalizeDocumentOoxml(volatileOoxml))
  assert.equal(hashDocumentOoxml(stableOoxml), hashDocumentOoxml(volatileOoxml))
} finally {
  globalThis.Word = originalWord
}

console.log('document map runtime tests: PASS')
