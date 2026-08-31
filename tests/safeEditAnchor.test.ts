import assert from 'node:assert/strict'

import { applyTextChangeProposal, makeSelectionProposal, restoreTextChange } from '../src/api/safeEdit.ts'

const originalWord = globalThis.Word

try {
  const makeRange = (initialText: string) => {
    const range = {
      text: initialText,
      ooxml: `<w:p><w:r><w:t>${initialText}</w:t></w:r></w:p>`,
      load: () => undefined,
      getOoxml: () => ({ value: range.ooxml }),
      insertContentControl: () => {
        const control = {
          tag: '',
          title: '',
          appearance: '',
          getRange: () => contentRange,
          delete: () => {
            controls.items = controls.items.filter(item => item !== control)
          },
        }
        const contentRange = {
          text: range.text,
          load: () => undefined,
          getOoxml: () => ({ value: range.ooxml }),
          insertText: (text: string) => {
            range.text = text
            range.ooxml = `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`
            contentRange.text = text
          },
          insertOoxml: (ooxml: string) => {
            range.ooxml = ooxml
            contentRange.text = ooxml.match(/<w:t>(.*?)<\/w:t>/)?.[1] || ''
            range.text = contentRange.text
          },
        }
        controls.items.push(control)
        return control
      },
    }
    return range
  }

  const originalRange = makeRange('original')
  const otherRange = makeRange('other')
  let selectedRange = originalRange
  const controls = { items: [] as any[], load: () => undefined }

  globalThis.Word = {
    run: async callback =>
      callback({
        document: {
          getSelection: () => selectedRange,
          contentControls: controls,
        },
        sync: async () => undefined,
      } as unknown as Word.RequestContext),
  } as unknown as typeof Word

  const proposal = await makeSelectionProposal('replacement', 'agent')
  selectedRange = otherRange
  const applied = await applyTextChangeProposal(proposal)

  assert.equal(originalRange.text, 'replacement')
  assert.equal(otherRange.text, 'other')
  await restoreTextChange(applied)
  assert.equal(originalRange.text, 'original')
  assert.equal(originalRange.ooxml, '<w:p><w:r><w:t>original</w:t></w:r></w:p>')
  assert.equal(controls.items.length, 0)
} finally {
  globalThis.Word = originalWord
}

console.log('safe edit anchor tests: PASS')
