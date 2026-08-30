import assert from 'node:assert/strict'

import { fetchModelCatalog, ModelCatalogError } from '../src/api/modelCatalog.ts'

const originalFetch = globalThis.fetch

try {
  let requestedURL = ''
  let requestedInit: RequestInit | undefined
  globalThis.fetch = async (input, init) => {
    requestedURL = String(input)
    requestedInit = init
    return new Response(JSON.stringify({ data: [{ id: 'model-10' }, { id: 'model-2' }, { id: 'model-2' }] }), {
      status: 200,
      headers: { 'content-type': 'application/json' },
    })
  }

  const openAIModels = await fetchModelCatalog({
    provider: 'official',
    apiKey: 'secret',
    baseURL: 'https://api.example.com/v1/',
    headers: { 'X-Project': 'word' },
  })
  assert.deepEqual(openAIModels, ['model-2', 'model-10'])
  assert.equal(requestedURL, 'https://api.example.com/v1/models')
  assert.deepEqual(requestedInit?.headers, { Authorization: 'Bearer secret', 'X-Project': 'word' })

  globalThis.fetch = async input => {
    requestedURL = String(input)
    return new Response(
      JSON.stringify({
        models: [
          { name: 'models/gemini-chat', supportedGenerationMethods: ['generateContent'] },
          { name: 'models/embedding-001', supportedGenerationMethods: ['embedContent'] },
        ],
      }),
      { status: 200, headers: { 'content-type': 'application/json' } },
    )
  }

  const geminiModels = await fetchModelCatalog({ provider: 'gemini', apiKey: 'gemini-secret' })
  assert.deepEqual(geminiModels, ['gemini-chat'])
  assert.match(requestedURL, /key=gemini-secret/)

  globalThis.fetch = async () =>
    new Response(JSON.stringify({ error: { message: 'Invalid key' } }), {
      status: 401,
      statusText: 'Unauthorized',
      headers: { 'content-type': 'application/json' },
    })

  await assert.rejects(
    () => fetchModelCatalog({ provider: 'official', apiKey: 'bad-key' }),
    (error: unknown) => error instanceof ModelCatalogError && error.status === 401 && error.message === 'Invalid key',
  )
} finally {
  globalThis.fetch = originalFetch
}

console.log('modelCatalog tests: PASS')
