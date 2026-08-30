export type ModelCatalogProvider = 'official' | 'gemini' | 'ollama' | 'groq'

export interface FetchModelCatalogOptions {
  provider: ModelCatalogProvider
  apiKey?: string
  baseURL?: string
  ollamaEndpoint?: string
  signal?: AbortSignal
}

const DEFAULT_ENDPOINTS = {
  official: 'https://api.openai.com/v1',
  groq: 'https://api.groq.com/openai/v1',
  ollama: 'http://localhost:11434',
} as const

const trimTrailingSlash = (value: string) => value.replace(/\/+$/, '')

const getErrorMessage = async (response: Response): Promise<string> => {
  try {
    const body = await response.json()
    return body?.error?.message || body?.message || `${response.status} ${response.statusText}`
  } catch {
    return `${response.status} ${response.statusText}`
  }
}

const requestJson = async (url: string, init: RequestInit = {}) => {
  const response = await fetch(url, init)
  if (!response.ok) {
    throw new Error(await getErrorMessage(response))
  }
  return response.json()
}

const requireApiKey = (apiKey: string | undefined) => {
  const value = apiKey?.trim()
  if (!value) throw new Error('API_KEY_REQUIRED')
  return value
}

const fetchOpenAICompatibleModels = async (
  baseURL: string,
  apiKey: string,
  signal?: AbortSignal,
): Promise<string[]> => {
  const data = await requestJson(`${trimTrailingSlash(baseURL)}/models`, {
    headers: { Authorization: `Bearer ${apiKey}` },
    signal,
  })

  if (!Array.isArray(data?.data)) throw new Error('INVALID_MODEL_RESPONSE')
  return data.data
    .map((model: { id?: string }) => model.id)
    .filter((id: unknown): id is string => typeof id === 'string')
}

const fetchGeminiModels = async (apiKey: string, signal?: AbortSignal): Promise<string[]> => {
  const models: string[] = []
  let pageToken = ''

  do {
    const url = new URL('https://generativelanguage.googleapis.com/v1beta/models')
    url.searchParams.set('key', apiKey)
    url.searchParams.set('pageSize', '1000')
    if (pageToken) url.searchParams.set('pageToken', pageToken)

    const data = await requestJson(url.toString(), { signal })
    if (!Array.isArray(data?.models)) throw new Error('INVALID_MODEL_RESPONSE')

    for (const model of data.models) {
      const methods = model?.supportedGenerationMethods
      if (Array.isArray(methods) && !methods.includes('generateContent')) continue
      if (typeof model?.name === 'string') models.push(model.name.replace(/^models\//, ''))
    }
    pageToken = typeof data.nextPageToken === 'string' ? data.nextPageToken : ''
  } while (pageToken)

  return models
}

const fetchOllamaModels = async (endpoint: string, signal?: AbortSignal): Promise<string[]> => {
  const data = await requestJson(`${trimTrailingSlash(endpoint)}/api/tags`, { signal })
  if (!Array.isArray(data?.models)) throw new Error('INVALID_MODEL_RESPONSE')
  return data.models
    .map((model: { name?: string; model?: string }) => model.name || model.model)
    .filter((name: unknown): name is string => typeof name === 'string')
}

export const fetchModelCatalog = async (options: FetchModelCatalogOptions): Promise<string[]> => {
  let models: string[]

  switch (options.provider) {
    case 'official':
      models = await fetchOpenAICompatibleModels(
        options.baseURL?.trim() || DEFAULT_ENDPOINTS.official,
        requireApiKey(options.apiKey),
        options.signal,
      )
      break
    case 'groq':
      models = await fetchOpenAICompatibleModels(DEFAULT_ENDPOINTS.groq, requireApiKey(options.apiKey), options.signal)
      break
    case 'gemini':
      models = await fetchGeminiModels(requireApiKey(options.apiKey), options.signal)
      break
    case 'ollama':
      models = await fetchOllamaModels(options.ollamaEndpoint?.trim() || DEFAULT_ENDPOINTS.ollama, options.signal)
      break
  }

  return [...new Set(models)].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
}
