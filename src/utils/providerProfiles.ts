import { computed, ref } from 'vue'

export type CapabilityState = 'yes' | 'no' | 'unknown'
export type ModelCapability = 'vision' | 'reasoning' | 'tools' | 'structuredOutput'

export interface ModelCapabilities {
  vision: CapabilityState
  reasoning: CapabilityState
  tools: CapabilityState
  structuredOutput: CapabilityState
}

export interface ProviderModel {
  id: string
  enabled: boolean
  capabilities: ModelCapabilities
  capabilityOverrides?: Partial<Record<ModelCapability, boolean>>
  lastVerifiedAt?: string
}

export interface CustomHeader {
  id: string
  key: string
  value: string
}

export interface ProviderProfile {
  id: string
  name: string
  protocol: 'openai-compatible'
  baseURL: string
  apiKey: string
  headers: CustomHeader[]
  defaultModel: string
  models: ProviderModel[]
  temperature: number
  maxTokens: number
  timeoutMs: number
  agentMaxIterations: number
  lastCheckedAt?: string
  lastConnectionStatus?: 'success' | 'failed'
}

export interface ProviderTemplate {
  id: string
  name: string
  baseURL: string
  defaultModel?: string
}

export const providerTemplates: ProviderTemplate[] = [
  { id: 'openai', name: 'OpenAI', baseURL: 'https://api.openai.com/v1', defaultModel: 'gpt-5' },
  { id: 'openrouter', name: 'OpenRouter', baseURL: 'https://openrouter.ai/api/v1' },
  { id: 'deepseek', name: 'DeepSeek', baseURL: 'https://api.deepseek.com/v1', defaultModel: 'deepseek-chat' },
  { id: 'siliconflow', name: 'SiliconFlow', baseURL: 'https://api.siliconflow.cn/v1' },
  {
    id: 'dashscope',
    name: 'Alibaba Bailian',
    baseURL: 'https://dashscope.aliyuncs.com/compatible-mode/v1',
  },
  { id: 'moonshot', name: 'Moonshot', baseURL: 'https://api.moonshot.cn/v1' },
  { id: 'custom', name: 'Custom', baseURL: '' },
]

const PROFILES_KEY = 'providerProfilesV1'
const ACTIVE_PROFILE_KEY = 'activeProviderProfileId'
const LEGACY_MIGRATION_KEY = 'providerProfilesLegacyMigrated'

const isNonChatModel = (id: string) =>
  /(embedding|embed-|rerank|moderation|whisper|tts|speech|transcri|image|dall-e|realtime|audio)/i.test(id)

export const inferModelCapabilities = (id: string): ModelCapabilities => {
  const name = id.toLowerCase()
  if (isNonChatModel(name)) {
    return { vision: 'no', reasoning: 'no', tools: 'no', structuredOutput: 'no' }
  }

  const vision =
    /(gpt-(4o|4\.1|5)|o[134]|claude-3|claude-4|gemini|qwen.*vl|vision|pixtral|llama.*vision|grok.*vision)/i.test(name)
  const reasoning =
    /(reason|thinking|deepseek-r1|deepseek-reasoner|qwq|qwen3|gpt-5|(^|\/)o[134]($|-)|gemini-(2\.5|3))/i.test(name)
  const tools =
    /(gpt-|(^|\/)o[134]($|-)|claude|gemini|deepseek-(chat|reasoner)|qwen|llama-(3\.1|3\.3|4)|mistral|mixtral|command-r|grok)/i.test(
      name,
    )
  const structured = /(gpt-|(^|\/)o[134]($|-)|claude|gemini|deepseek|qwen|llama-(3\.1|3\.3|4)|mistral|grok)/i.test(name)

  return {
    vision: vision ? 'yes' : 'unknown',
    reasoning: reasoning ? 'yes' : 'unknown',
    tools: tools ? 'yes' : 'unknown',
    structuredOutput: structured ? 'yes' : 'unknown',
  }
}

export const createProviderModel = (id: string, enabled = true): ProviderModel => ({
  id,
  enabled,
  capabilities: inferModelCapabilities(id),
})

const getLegacyModels = (): string[] => {
  const stored = localStorage.getItem('customModels')
  if (stored) {
    try {
      const parsed = JSON.parse(stored)
      if (Array.isArray(parsed)) return parsed.filter((item): item is string => typeof item === 'string')
    } catch {
      // Ignore malformed legacy data.
    }
  }
  const selected = localStorage.getItem('model')
  return selected ? [selected] : []
}

const createProfileFromTemplate = (template: ProviderTemplate): ProviderProfile => {
  const id = crypto.randomUUID()
  const models = template.defaultModel ? [createProviderModel(template.defaultModel)] : []
  return {
    id,
    name: template.name,
    protocol: 'openai-compatible',
    baseURL: template.baseURL,
    apiKey: '',
    headers: [],
    defaultModel: template.defaultModel || '',
    models,
    temperature: 0.7,
    maxTokens: 0,
    timeoutMs: 60000,
    agentMaxIterations: 25,
  }
}

const createLegacyProfile = (): ProviderProfile => {
  const baseURL = localStorage.getItem('basePath') || 'https://api.openai.com/v1'
  const selectedModel = localStorage.getItem('model') || 'gpt-5'
  const modelIds = [...new Set([selectedModel, ...getLegacyModels()])]
  const isOfficial = /api\.openai\.com/i.test(baseURL)
  return {
    id: crypto.randomUUID(),
    name: isOfficial ? 'OpenAI' : 'Migrated Provider',
    protocol: 'openai-compatible',
    baseURL,
    apiKey: localStorage.getItem('apiKey') || '',
    headers: [],
    defaultModel: selectedModel,
    models: modelIds.map(model => createProviderModel(model)),
    temperature: Number(localStorage.getItem('temperature') || 0.7),
    maxTokens:
      Number(localStorage.getItem('maxTokens') || 0) === 800 ? 0 : Number(localStorage.getItem('maxTokens') || 0),
    timeoutMs: 60000,
    agentMaxIterations: Number(localStorage.getItem('agentMaxIterations') || 25),
  }
}

const normalizeProfile = (profile: Partial<ProviderProfile>): ProviderProfile => {
  const template = createProviderProfile('openai')
  const models = Array.isArray(profile.models)
    ? profile.models
        .filter(model => model && typeof model.id === 'string')
        .map(model => ({
          ...createProviderModel(model.id, model.enabled !== false),
          ...model,
          capabilities: { ...inferModelCapabilities(model.id), ...(model.capabilities || {}) },
        }))
    : []
  return {
    ...template,
    ...profile,
    id: profile.id || crypto.randomUUID(),
    headers: Array.isArray(profile.headers) ? profile.headers : [],
    models,
  }
}

function createProviderProfile(templateId: string): ProviderProfile {
  const template = providerTemplates.find(item => item.id === templateId) || providerTemplates.at(-1)!
  return createProfileFromTemplate(template)
}

const loadProfiles = (): ProviderProfile[] => {
  const stored = localStorage.getItem(PROFILES_KEY)
  if (stored) {
    try {
      const parsed = JSON.parse(stored)
      if (Array.isArray(parsed) && parsed.length > 0) return parsed.map(normalizeProfile)
    } catch {
      // Fall through to legacy migration.
    }
  }
  const migrated = createLegacyProfile()
  localStorage.setItem(LEGACY_MIGRATION_KEY, '1')
  localStorage.setItem(PROFILES_KEY, JSON.stringify([migrated]))
  return [migrated]
}

const profiles = ref<ProviderProfile[]>(loadProfiles())
const storedActiveId = localStorage.getItem(ACTIVE_PROFILE_KEY)
const activeProfileId = ref(
  profiles.value.some(profile => profile.id === storedActiveId) ? storedActiveId! : profiles.value[0]?.id || '',
)

const persistProfiles = () => localStorage.setItem(PROFILES_KEY, JSON.stringify(profiles.value))

export const cloneProviderProfile = (profile: ProviderProfile): ProviderProfile => JSON.parse(JSON.stringify(profile))

export const useProviderProfiles = () => {
  const activeProfile = computed(() => profiles.value.find(profile => profile.id === activeProfileId.value) || null)

  const setActiveProfile = (id: string) => {
    if (!profiles.value.some(profile => profile.id === id)) return
    activeProfileId.value = id
    localStorage.setItem(ACTIVE_PROFILE_KEY, id)
  }

  const addProfile = (templateId: string) => {
    const profile = createProviderProfile(templateId)
    profiles.value.push(profile)
    persistProfiles()
    setActiveProfile(profile.id)
    return profile
  }

  const updateProfile = (profile: ProviderProfile) => {
    const index = profiles.value.findIndex(item => item.id === profile.id)
    if (index === -1) return
    profiles.value[index] = normalizeProfile(cloneProviderProfile(profile))
    persistProfiles()
  }

  const removeProfile = (id: string) => {
    if (profiles.value.length <= 1) return false
    if (!profiles.value.some(profile => profile.id === id)) return false
    profiles.value = profiles.value.filter(profile => profile.id !== id)
    if (activeProfileId.value === id) setActiveProfile(profiles.value[0].id)
    persistProfiles()
    return true
  }

  return {
    profiles,
    activeProfileId,
    activeProfile,
    setActiveProfile,
    addProfile,
    updateProfile,
    removeProfile,
  }
}

export const getResolvedCapability = (model: ProviderModel, capability: ModelCapability): CapabilityState => {
  const override = model.capabilityOverrides?.[capability]
  return override === undefined ? model.capabilities[capability] : override ? 'yes' : 'no'
}

export const cycleCapabilityOverride = (model: ProviderModel, capability: ModelCapability) => {
  model.capabilityOverrides ||= {}
  const current = model.capabilityOverrides[capability]
  if (current === undefined) model.capabilityOverrides[capability] = true
  else if (current === true) model.capabilityOverrides[capability] = false
  else delete model.capabilityOverrides[capability]
}

export const resolveMaxTokens = (configured: number, model: ProviderModel | undefined, agentMode: boolean) => {
  if (configured > 0) return configured
  const reasoning = model ? getResolvedCapability(model, 'reasoning') === 'yes' : false
  if (reasoning) return agentMode ? 16384 : 8192
  return agentMode ? 8192 : 4096
}
