import assert from 'node:assert/strict'

import { reactive } from 'vue'

class MemoryStorage implements Storage {
  private values = new Map<string, string>()

  get length() {
    return this.values.size
  }

  clear() {
    this.values.clear()
  }

  getItem(key: string) {
    return this.values.get(key) ?? null
  }

  key(index: number) {
    return [...this.values.keys()][index] ?? null
  }

  removeItem(key: string) {
    this.values.delete(key)
  }

  setItem(key: string, value: string) {
    this.values.set(key, String(value))
  }
}

globalThis.localStorage = new MemoryStorage()
localStorage.setItem('apiKey', 'legacy-key')
localStorage.setItem('basePath', 'https://api.example.com/v1')
localStorage.setItem('model', 'gpt-5')
localStorage.setItem('maxTokens', '800')

const {
  cloneProviderProfile,
  createProviderModel,
  cycleCapabilityOverride,
  getResolvedCapability,
  inferModelCapabilities,
  resolveMaxTokens,
  useProviderProfiles,
} = await import('../src/utils/providerProfiles.ts')

const embedding = inferModelCapabilities('text-embedding-3-large')
assert.equal(embedding.tools, 'no')
assert.equal(embedding.vision, 'no')

const reasoning = createProviderModel('gpt-5')
assert.equal(getResolvedCapability(reasoning, 'tools'), 'yes')
assert.equal(getResolvedCapability(reasoning, 'reasoning'), 'yes')
assert.equal(resolveMaxTokens(0, reasoning, false), 8192)
assert.equal(resolveMaxTokens(0, reasoning, true), 16384)
assert.equal(resolveMaxTokens(1234, reasoning, true), 1234)

const unknown = createProviderModel('private-company-model')
assert.equal(getResolvedCapability(unknown, 'tools'), 'unknown')
assert.equal(resolveMaxTokens(0, unknown, false), 4096)
assert.equal(resolveMaxTokens(0, unknown, true), 8192)

cycleCapabilityOverride(unknown, 'tools')
assert.equal(getResolvedCapability(unknown, 'tools'), 'yes')
cycleCapabilityOverride(unknown, 'tools')
assert.equal(getResolvedCapability(unknown, 'tools'), 'no')
cycleCapabilityOverride(unknown, 'tools')
assert.equal(getResolvedCapability(unknown, 'tools'), 'unknown')

const { profiles, activeProfile, activeProfileId, addProfile, removeProfile, updateProfile } = useProviderProfiles()
assert.equal(profiles.value.length, 1)
assert.equal(activeProfile.value?.apiKey, 'legacy-key')
assert.equal(activeProfile.value?.baseURL, 'https://api.example.com/v1')
assert.equal(activeProfile.value?.defaultModel, 'gpt-5')
assert.equal(activeProfile.value?.maxTokens, 0)

const reactiveProfile = reactive(activeProfile.value!)
const clonedProfile = cloneProviderProfile(reactiveProfile)
assert.notEqual(clonedProfile, reactiveProfile)
assert.equal(clonedProfile.apiKey, 'legacy-key')
assert.deepEqual(clonedProfile.models, reactiveProfile.models)

const originalProfileId = activeProfileId.value
const cpaProfile = addProfile('custom')
cpaProfile.name = 'cpa'
updateProfile(cpaProfile)
assert.equal(profiles.value.length, 2)
assert.equal(activeProfile.value?.name, 'cpa')
assert.equal(removeProfile(cpaProfile.id), true)
assert.equal(profiles.value.length, 1)
assert.equal(activeProfileId.value, originalProfileId)
assert.equal(localStorage.getItem('activeProviderProfileId'), originalProfileId)
assert.equal(removeProfile('missing-profile'), false)
assert.equal(removeProfile(originalProfileId), false)
assert.equal(profiles.value.length, 1)

console.log('providerProfiles tests: PASS')
