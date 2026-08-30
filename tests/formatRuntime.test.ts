import assert from 'node:assert/strict'

import { readSelectionInspection } from '../src/api/safeEdit.ts'
import { createTaskTools, createTaskToolState } from '../src/utils/taskTools.ts'

const originalWord = globalThis.Word

const installWordMock = (run: (callback: (context: Word.RequestContext) => unknown) => Promise<unknown>) => {
  globalThis.Word = { run } as unknown as typeof Word
}

try {
  // Read-only inspection falls back to text when selection.getOoxml() fails.
  let inspectionRuns = 0
  installWordMock(async callback => {
    inspectionRuns++
    const shouldFailOoxmlSync = inspectionRuns === 1
    const range = {
      text: 'text across a page break\f',
      getOoxml: () => ({ value: '' }),
      load: () => undefined,
    }
    const context = {
      document: { getSelection: () => range },
      sync: async () => {
        if (shouldFailOoxmlSync) throw new Error('GeneralException')
      },
    }
    return callback(context as unknown as Word.RequestContext)
  })

  const inspection = await readSelectionInspection()
  assert.equal(inspection.text, 'text across a page break\f')
  assert.equal(inspection.protectedObjectsAvailable, false)
  assert.equal(inspection.protectedObjects.length, 0)
  assert.match(inspection.warning || '', /could not inspect non-text objects/i)
  assert.equal(inspectionRuns, 2)

  // Formatting proposals and application use the text hash only, so OOXML is
  // never requested. Applying an active proposal does not require window.confirm.
  let ooxmlCalls = 0
  const paragraph = {
    alignment: 'Left',
    lineSpacing: 12,
    spaceAfter: 0,
    load: () => undefined,
  }
  const range = {
    text: 'influence ',
    load: () => undefined,
    getOoxml: () => {
      ooxmlCalls++
      throw new Error('getOoxml should not be called by formatting')
    },
    font: {
      name: 'Calibri',
      size: 11,
      bold: false,
      italic: false,
      underline: 'None',
      color: '#000000',
      highlightColor: '',
      load: () => undefined,
    },
    paragraphs: {
      items: [paragraph],
      load: () => undefined,
    },
  }
  installWordMock(async callback => {
    const context = {
      document: {
        getSelection: () => range,
        body: { getRange: () => range },
      },
      sync: async () => undefined,
    }
    return callback(context as unknown as Word.RequestContext)
  })

  const state = createTaskToolState()
  const proposeTool = createTaskTools(['propose_format_patch'], undefined, state)[0]
  const proposal = JSON.parse(String(await proposeTool.invoke({ scope: 'selection', fontSize: 12 })))
  assert.deepEqual(proposal.changes, { fontSize: 12 })
  assert.equal(state.activeFormatRequestId, proposal.formatId)

  const applyTool = createTaskTools(['apply_format_patch'], undefined, state)[0]
  const applied = JSON.parse(String(await applyTool.invoke({ formatId: proposal.formatId })))
  assert.equal(applied.status, 'applied')
  assert.deepEqual(applied.changes, { fontSize: 12 })
  assert.equal(applied.verification.verified, true)
  assert.equal(range.font.size, 12)
  assert.equal(ooxmlCalls, 0)
  assert.equal(state.activeFormatRequestId, null)

  const directTool = createTaskTools(['format_document_selection'], { allowedFormatFields: ['fontSize'] }, state)[0]
  const directResult = JSON.parse(
    String(await directTool.invoke({ scope: 'selection', fontSize: 14, fontName: 'Arial', bold: true })),
  )
  assert.equal(directResult.status, 'applied')
  assert.deepEqual(directResult.changes, { fontSize: 14 })
  assert.equal(directResult.verification.verified, true)
  assert.equal(range.font.size, 14)
  assert.equal(range.font.name, 'Calibri')
  assert.equal(range.font.bold, false)
} finally {
  globalThis.Word = originalWord
}

console.log('format runtime tests: PASS')
