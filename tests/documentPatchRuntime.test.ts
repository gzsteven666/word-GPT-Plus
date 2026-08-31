import assert from 'node:assert/strict'

import {
  applyDocumentPatchSet,
  proposeDocumentPatchSet,
  restoreDocumentPatchSet,
} from '../src/api/documentPatchRuntime.ts'
import { acceptPatchOperation } from '../src/utils/documentPatch.ts'

const originalWord = globalThis.Word

try {
  const controls = { items: [] as any[], load: () => undefined }
  const bodyText = '第一处\n第二处'
  let selectedFailure = false
  const makeRange = (text: string) => {
    const range: any = {
      text,
      ooxml: `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`,
      load: () => undefined,
      getOoxml: () => ({ value: range.ooxml }),
      compareLocationWith: () => ({ value: 'Before' }),
      insertContentControl: () => {
        const contentRange: any = {
          text: range.text,
          load: () => undefined,
          getOoxml: () => ({ value: range.ooxml }),
          insertText: (value: string, location: string) => {
            if (selectedFailure && range.text === '第一处') throw new Error('simulated second write failure')
            if (location === 'Replace') range.text = value
            else if (location === 'Before') range.text = `${value}${range.text}`
            else range.text = `${range.text}${value}`
            contentRange.text = range.text
            range.ooxml = `<w:p><w:r><w:t>${range.text}</w:t></w:r></w:p>`
          },
          insertOoxml: (value: string) => {
            range.ooxml = value
            range.text = value.match(/<w:t>(.*?)<\/w:t>/)?.[1] || ''
            contentRange.text = range.text
          },
        }
        const control: any = {
          tag: '',
          title: '',
          appearance: '',
          getRange: () => contentRange,
          delete: () => {
            controls.items = controls.items.filter(item => item !== control)
          },
        }
        controls.items.push(control)
        return control
      },
    }
    return range
  }

  const ranges = [makeRange('第一处'), makeRange('第二处')]
  const duplicateRanges = [makeRange('重复'), makeRange('重复')]
  const overlapRanges = [makeRange('重叠一'), makeRange('重叠二')]
  overlapRanges[0].compareLocationWith = () => ({ value: 'Overlap' })
  const context = {
    document: {
      body: {
        search: (query: string) => ({
          items:
            query === '重复'
              ? duplicateRanges
              : query === '重叠一'
                ? [overlapRanges[0]]
                : query === '重叠二'
                  ? [overlapRanges[1]]
                  : ranges.filter(range => range.text === query),
          load: () => undefined,
        }),
        getRange: () => ({ text: bodyText, getOoxml: () => ({ value: '<w:body />' }), load: () => undefined }),
      },
      contentControls: controls,
    },
    sync: async () => undefined,
  }
  globalThis.Word = {
    run: async callback => callback(context as unknown as Word.RequestContext),
  } as unknown as typeof Word

  const proposed = await proposeDocumentPatchSet([
    { type: 'replace', targetText: '第一处', replacementText: '修改一' },
    { type: 'replace', targetText: '第二处', replacementText: '修改二' },
  ])
  assert.equal(proposed.operations.length, 2)
  assert.equal(controls.items.length, 2)
  await assert.rejects(
    () => proposeDocumentPatchSet([{ type: 'replace', targetText: '重复', replacementText: '不应自动选择' }]),
    /ambiguous/i,
  )
  await assert.rejects(
    () =>
      proposeDocumentPatchSet([
        { type: 'replace', targetText: '重叠一', replacementText: '一' },
        { type: 'replace', targetText: '重叠二', replacementText: '二' },
      ]),
    /overlap/i,
  )
  acceptPatchOperation(proposed, proposed.operations[0].id)
  acceptPatchOperation(proposed, proposed.operations[1].id)

  const conflict = await proposeDocumentPatchSet([{ type: 'replace', targetText: '第一处', replacementText: '冲突' }])
  acceptPatchOperation(conflict, conflict.operations[0].id)
  const conflictControl = controls.items.find(item => item.tag === conflict.operations[0].anchorTag)
  conflictControl.getRange().text = '人工修改'
  await assert.rejects(() => applyDocumentPatchSet(conflict), /changed/i)
  conflictControl.getRange().text = '第一处'
  ranges[0].ooxml = '<w:p><w:r><w:t>第一处</w:t></w:r></w:p>'

  selectedFailure = true
  await assert.rejects(() => applyDocumentPatchSet(proposed), /simulated second write failure|rolled back/i)
  assert.equal(ranges[0].text, '第一处')
  assert.equal(ranges[1].text, '第二处')

  selectedFailure = false
  const retry = await proposeDocumentPatchSet([
    { type: 'replace', targetText: '第一处', replacementText: '修改一' },
    { type: 'replace', targetText: '第二处', replacementText: '修改二' },
  ])
  acceptPatchOperation(retry, retry.operations[0].id)
  acceptPatchOperation(retry, retry.operations[1].id)
  const applied = await applyDocumentPatchSet(retry)
  assert.equal(applied.status, 'applied')
  const restored = await restoreDocumentPatchSet(applied)
  assert.equal(restored.status, 'restored')
  assert.equal(ranges[0].text, '第一处')
  assert.equal(bodyText, '第一处\n第二处')
} finally {
  globalThis.Word = originalWord
}

console.log('document patch runtime tests: PASS')
