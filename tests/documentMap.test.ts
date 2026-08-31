import assert from 'node:assert/strict'

import {
  buildDocumentMap,
  hasDocumentMapTargetReference,
  isDocumentMapCurrent,
  isDocumentMapIntent,
  queryDocumentMap,
  selectDocumentMapNodes,
} from '../src/utils/documentMap.ts'

const input = {
  documentHash: 'doc-hash-1',
  documentOoxmlHash: 'doc-ooxml-1',
  paragraphs: [
    { text: '项目报告', styleBuiltIn: 'Heading1' },
    { text: '收入同比增长 20%。' },
    { text: '风险与计划', styleBuiltIn: 'Heading2' },
    { text: '下一步完成验收。' },
  ],
}

const map = buildDocumentMap(input)
assert.equal(map.nodes.length, 4)
assert.deepEqual(map.nodes[0].headingPath, ['项目报告'])
assert.deepEqual(map.nodes[1].headingPath, ['项目报告'])
assert.deepEqual(map.nodes[2].headingPath, ['项目报告', '风险与计划'])
assert.equal(map.nodes[0].headingLevel, 1)
assert.equal(map.nodes[2].headingLevel, 2)
assert.equal(map.nodes[1].preview, '收入同比增长 20%。')
assert.equal(buildDocumentMap(input).nodes[1].id, map.nodes[1].id)

const results = queryDocumentMap(map, '风险 计划')
assert.equal(results.length, 2)
assert.equal(results[0].node.id, map.nodes[2].id)
assert.ok(results[0].score > 0)

const selected = selectDocumentMapNodes(map, [map.nodes[1].id, map.nodes[3].id], 12)
assert.equal(selected.length, 2)
assert.equal(selected[0].node.id, map.nodes[1].id)
assert.equal(selected.reduce((total, item) => total + item.text.length, 0) <= 12, true)

assert.equal(isDocumentMapCurrent(map, 'doc-hash-1', 'doc-ooxml-1'), true)
assert.equal(isDocumentMapCurrent(map, 'doc-hash-2', 'doc-ooxml-1'), false)
assert.equal(isDocumentMapCurrent(map, 'doc-hash-1', 'doc-ooxml-2'), false)
assert.equal(isDocumentMapIntent('请使用刚才的新 mapId 查询目标文本'), true)
assert.equal(isDocumentMapIntent('请调整选中文本字号'), false)
assert.equal(hasDocumentMapTargetReference('使用 mapId map-e40cbdc4 和 nodeId node-4-cd475760 替换目标文本'), true)
assert.equal(hasDocumentMapTargetReference('使用 mapId map-e40cbdc4 查询目标文本'), false)

console.log('document map tests: PASS')
