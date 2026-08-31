import assert from 'node:assert/strict'

import { readSelectionInspection } from '../src/api/safeEdit.ts'
import { createTaskTools, createTaskToolState, restoreFormatRequest } from '../src/utils/taskTools.ts'

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

  // Formatting proposals keep a hidden anchor and verify both text and OOXML.
  let ooxmlCalls = 0
  const controls = { items: [] as any[], load: () => undefined }
  const paragraph = {
    alignment: 'Left',
    lineSpacing: 12,
    spaceAfter: 0,
    load: () => undefined,
  }
  const range = {
    text: 'influence ',
    ooxml: '<w:p><w:r><w:t>influence </w:t></w:r></w:p>',
    load: () => undefined,
    getOoxml: () => {
      ooxmlCalls++
      return { value: range.ooxml }
    },
    insertContentControl: () => {
      const control = {
        tag: '',
        title: '',
        appearance: '',
        getRange: () => range,
        delete: () => undefined,
      }
      controls.items.push(control)
      return control
    },
    insertOoxml: (ooxml: string) => {
      range.ooxml = ooxml
      range.font.size = 11
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
  const otherRange = {
    ...range,
    font: { ...range.font },
    getOoxml: () => ({ value: '<w:p><w:r><w:t>other</w:t></w:r></w:p>' }),
  }
  otherRange.insertContentControl = () => {
    const control = {
      tag: '',
      title: '',
      appearance: '',
      getRange: () => otherRange,
      delete: () => undefined,
    }
    controls.items.push(control)
    return control
  }
  let currentSelection = range
  installWordMock(async callback => {
    const context = {
      document: {
        getSelection: () => currentSelection,
        contentControls: controls,
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
  currentSelection = otherRange

  const applyTool = createTaskTools(['apply_format_patch'], undefined, state)[0]
  const applied = JSON.parse(String(await applyTool.invoke({ formatId: proposal.formatId })))
  assert.equal(applied.status, 'applied')
  assert.deepEqual(applied.changes, { fontSize: 12 })
  assert.equal(applied.verification.verified, true)
  assert.equal(range.font.size, 12)
  assert.equal(otherRange.font.size, 11)
  assert.ok(ooxmlCalls > 0)
  assert.equal(state.activeFormatRequestId, null)
  await restoreFormatRequest(state.appliedFormatRequests.get(proposal.formatId)!)
  assert.equal(range.font.size, 11)

  const directTool = createTaskTools(['format_document_selection'], { allowedFormatFields: ['fontSize'] }, state)[0]
  const directResult = JSON.parse(
    String(await directTool.invoke({ scope: 'selection', fontSize: 14, fontName: 'Arial', bold: true })),
  )
  assert.equal(directResult.status, 'applied')
  assert.deepEqual(directResult.changes, { fontSize: 14 })
  assert.equal(directResult.verification.verified, true)
  assert.equal(otherRange.font.size, 14)
  assert.equal(otherRange.font.name, 'Calibri')
  assert.equal(otherRange.font.bold, false)
} finally {
  globalThis.Word = originalWord
}

console.log('format runtime tests: PASS')
