<template>
  <div class="flex w-full flex-col gap-2">
    <SettingCard p1>
      <div class="flex flex-col gap-3 p-3">
        <div class="flex items-center justify-between gap-2">
          <div>
            <h3 class="text-sm font-semibold text-main">{{ t('providerProfiles') }}</h3>
            <p class="text-xs text-secondary">{{ t('providerProfilesDescription') }}</p>
          </div>
          <span class="rounded-full bg-accent/10 px-2 py-1 text-xs text-accent">{{ profiles.length }}</span>
        </div>

        <div class="flex gap-2 overflow-x-auto pb-1">
          <button
            v-for="profile in profiles"
            :key="profile.id"
            class="min-w-32 rounded-md border px-3 py-2 text-left transition-colors"
            :class="
              activeProfileId === profile.id
                ? 'border-accent bg-accent/10 text-main'
                : 'border-border bg-surface text-secondary hover:border-accent/50'
            "
            @click="selectProfile(profile.id)"
          >
            <span class="block truncate text-xs font-semibold">{{ profile.name }}</span>
            <span class="mt-1 block text-[10px]">
              {{ profile.lastConnectionStatus === 'success' ? t('connected') : t('notVerified') }} ·
              {{ profile.models.filter(model => model.enabled).length }} {{ t('models') }}
            </span>
          </button>
        </div>

        <div class="flex gap-2">
          <select
            v-model="selectedTemplate"
            class="min-w-0 flex-1 rounded-md border border-border bg-bg-tertiary p-2 text-xs text-main outline-none focus:border-accent"
          >
            <option v-for="template in providerTemplates" :key="template.id" :value="template.id">
              {{ template.name }}
            </option>
          </select>
          <CustomButton :icon="Plus" :text="t('addProvider')" type="secondary" @click="addProvider" />
        </div>
      </div>
    </SettingCard>

    <template v-if="draft">
      <SettingCard p1>
        <div class="flex flex-col gap-3 p-3">
          <div class="flex items-center justify-between gap-2">
            <h3 class="text-sm font-semibold text-main">{{ t('providerConfiguration') }}</h3>
            <button
              v-if="profiles.length > 1"
              class="rounded-md p-1.5 text-danger hover:bg-danger/10"
              :title="t('deleteProvider')"
              @click="deleteProvider"
            >
              <Trash2 :size="16" />
            </button>
          </div>

          <CustomInput v-model="draft.name" :title="t('providerName')" :placeholder="t('providerNamePlaceholder')" />
          <CustomInput v-model="draft.baseURL" :title="t('baseURL')" placeholder="https://api.example.com/v1" />
          <p class="-mt-2 text-[11px] break-all text-secondary">
            {{ t('modelsRequestPreview') }}: {{ modelsEndpoint }}
          </p>
          <CustomInput v-model="draft.apiKey" :title="t('apiKey')" :placeholder="t('apiKeyPlaceholder')" is-password />
          <p class="-mt-2 text-[11px] text-secondary">{{ t('apiKeyLocalNotice') }}</p>

          <details class="rounded-md border border-border bg-surface p-2">
            <summary class="cursor-pointer text-xs font-semibold text-secondary">{{ t('advancedSettings') }}</summary>
            <div class="mt-3 flex flex-col gap-3">
              <div class="grid grid-cols-2 gap-2">
                <CustomInput
                  v-model.number="draft.temperature"
                  :title="t('temperature')"
                  placeholder="0.7"
                  input-type="number"
                  min="0"
                  max="2"
                  step="0.1"
                />
                <CustomInput
                  v-model.number="draft.maxTokens"
                  :title="t('maxOutputTokens')"
                  :placeholder="t('zeroMeansAuto')"
                  input-type="number"
                  min="0"
                  max="131072"
                  step="1"
                />
                <CustomInput
                  v-model.number="draft.timeoutMs"
                  :title="t('requestTimeoutMs')"
                  placeholder="60000"
                  input-type="number"
                  min="1000"
                  max="600000"
                  step="1000"
                />
                <CustomInput
                  v-model.number="draft.agentMaxIterations"
                  :title="t('agentMaxIterationsLabel')"
                  placeholder="25"
                  input-type="number"
                  min="1"
                  max="500"
                  step="1"
                />
              </div>

              <div class="flex items-center justify-between">
                <span class="text-xs font-semibold text-secondary">{{ t('customHeaders') }}</span>
                <button class="text-xs text-accent hover:underline" @click="addHeader">+ {{ t('addHeader') }}</button>
              </div>
              <div v-for="header in draft.headers" :key="header.id" class="grid grid-cols-[1fr_1fr_auto] gap-2">
                <input
                  v-model="header.key"
                  class="min-w-0 rounded-md border border-border bg-bg-tertiary p-2 text-xs text-main outline-none focus:border-accent"
                  :placeholder="t('headerName')"
                />
                <input
                  v-model="header.value"
                  type="password"
                  class="min-w-0 rounded-md border border-border bg-bg-tertiary p-2 text-xs text-main outline-none focus:border-accent"
                  :placeholder="t('headerValue')"
                />
                <button class="rounded-md p-2 text-danger hover:bg-danger/10" @click="removeHeader(header.id)">
                  <X :size="14" />
                </button>
              </div>
            </div>
          </details>

          <div
            v-if="connectionMessage"
            class="rounded-md border p-2 text-xs"
            :class="
              connectionOk
                ? 'border-success/30 bg-success/10 text-success'
                : 'border-danger/30 bg-danger/10 text-danger'
            "
          >
            {{ connectionMessage }}
          </div>

          <div class="grid grid-cols-2 gap-2">
            <CustomButton
              :icon="testing ? LoaderCircle : PlugZap"
              :icon-class="testing ? 'animate-spin' : ''"
              :text="t('testConnection')"
              type="secondary"
              :disabled="testing || fetching"
              @click="testConnection"
            />
            <CustomButton
              :icon="fetching ? LoaderCircle : Download"
              :icon-class="fetching ? 'animate-spin' : ''"
              :text="t('fetchModels')"
              type="secondary"
              :disabled="testing || fetching"
              @click="fetchModels"
            />
          </div>
        </div>
      </SettingCard>

      <SettingCard p1>
        <div class="flex flex-col gap-3 p-3">
          <div class="flex items-center justify-between gap-2">
            <div>
              <h3 class="text-sm font-semibold text-main">{{ t('enabledModels') }}</h3>
              <p class="text-[11px] text-secondary">{{ t('capabilityOverrideHint') }}</p>
            </div>
            <span class="text-xs text-secondary">{{ enabledModels.length }}/{{ draft.models.length }}</span>
          </div>

          <div class="flex gap-2">
            <input
              v-model="manualModel"
              class="min-w-0 flex-1 rounded-md border border-border bg-bg-tertiary p-2 text-xs text-main outline-none focus:border-accent"
              :placeholder="t('customModelPlaceholder')"
              @keyup.enter="addManualModel"
            />
            <CustomButton :icon="Plus" text="" type="secondary" @click="addManualModel" />
          </div>

          <div v-if="draft.models.length" class="flex max-h-80 flex-col gap-1 overflow-y-auto">
            <div
              v-for="model in draft.models"
              :key="model.id"
              class="flex items-center gap-2 rounded-md border border-border bg-surface p-2"
            >
              <input v-model="model.enabled" type="checkbox" class="h-4 w-4 shrink-0 accent-accent" />
              <button
                class="min-w-0 flex-1 truncate text-left text-xs"
                :class="draft.defaultModel === model.id ? 'font-semibold text-accent' : 'text-main'"
                :title="model.id"
                @click="setDefaultModel(model.id)"
              >
                {{ model.id }}
                <span v-if="draft.defaultModel === model.id"> · {{ t('defaultModel') }}</span>
              </button>
              <div class="flex shrink-0 gap-1">
                <button
                  v-for="capability in capabilityItems"
                  :key="capability.key"
                  class="rounded-full p-1"
                  :class="capabilityClass(model, capability.key)"
                  :title="capabilityTitle(model, capability.key, capability.label)"
                  @click="cycleCapabilityOverride(model, capability.key)"
                >
                  <component :is="capability.icon" :size="12" />
                </button>
              </div>
              <button class="shrink-0 rounded-md p-1 text-danger hover:bg-danger/10" @click="removeModel(model.id)">
                <X :size="13" />
              </button>
            </div>
          </div>
          <p v-else class="py-4 text-center text-xs text-secondary">{{ t('noModelsConfigured') }}</p>

          <div class="flex justify-end gap-2 border-t border-border pt-3">
            <CustomButton :text="t('cancel')" type="secondary" :disabled="!dirty" @click="resetDraft" />
            <CustomButton :icon="Save" :text="t('save')" :disabled="!dirty" @click="saveProfile" />
          </div>
        </div>
      </SettingCard>
    </template>

    <ModelPickerDialog
      :open="pickerOpen"
      :models="catalogModels"
      :existing-models="draft?.models.map(model => model.id) || []"
      :title="t('modelPickerTitle')"
      :subtitle="t('modelPickerSubtitle', { provider: draft?.name || '' })"
      :search-placeholder="t('searchModels')"
      :close-label="t('cancel')"
      :select-visible-label="t('selectVisible')"
      :clear-visible-label="t('clearVisible')"
      :added-label="t('alreadyAdded')"
      :empty-label="t('noModelsFound')"
      :selected-label="t('modelsSelected')"
      :cancel-label="t('cancel')"
      :add-selected-label="t('addSelectedModels')"
      :result-summary="t('modelsFound', { count: catalogModels.length })"
      @close="pickerOpen = false"
      @confirm="addFetchedModels"
    />
  </div>
</template>

<script setup lang="ts">
import { Braces, Brain, Download, Eye, LoaderCircle, PlugZap, Plus, Save, Trash2, Wrench, X } from 'lucide-vue-next'
import { computed, ref, watch } from 'vue'
import { useI18n } from 'vue-i18n'

import { fetchModelCatalog, ModelCatalogError } from '@/api/modelCatalog'
import CustomButton from '@/components/CustomButton.vue'
import CustomInput from '@/components/CustomInput.vue'
import ModelPickerDialog from '@/components/ModelPickerDialog.vue'
import SettingCard from '@/components/SettingCard.vue'
import {
  createProviderModel,
  cycleCapabilityOverride,
  getResolvedCapability,
  ModelCapability,
  ProviderModel,
  ProviderProfile,
  providerTemplates,
  useProviderProfiles,
} from '@/utils/providerProfiles'

const { t } = useI18n()
const { profiles, activeProfileId, activeProfile, setActiveProfile, addProfile, updateProfile, removeProfile } =
  useProviderProfiles()

const selectedTemplate = ref('openai')
const draft = ref<ProviderProfile | null>(activeProfile.value ? structuredClone(activeProfile.value) : null)
const savedSnapshot = ref(draft.value ? JSON.stringify(draft.value) : '')
const testing = ref(false)
const fetching = ref(false)
const connectionMessage = ref('')
const connectionOk = ref(false)
const pickerOpen = ref(false)
const catalogModels = ref<string[]>([])
const manualModel = ref('')

const forbiddenHeaders = new Set(['host', 'origin', 'content-length', 'cookie'])
const capabilityItems: { key: ModelCapability; label: string; icon: any }[] = [
  { key: 'vision', label: 'visionCapability', icon: Eye },
  { key: 'reasoning', label: 'reasoningCapability', icon: Brain },
  { key: 'tools', label: 'toolCapability', icon: Wrench },
  { key: 'structuredOutput', label: 'structuredOutputCapability', icon: Braces },
]

const dirty = computed(() => !!draft.value && JSON.stringify(draft.value) !== savedSnapshot.value)
const enabledModels = computed(() => draft.value?.models.filter(model => model.enabled) || [])
const modelsEndpoint = computed(() => `${draft.value?.baseURL?.replace(/\/+$/, '') || '—'}/models`)

watch(
  activeProfile,
  profile => {
    draft.value = profile ? structuredClone(profile) : null
    savedSnapshot.value = draft.value ? JSON.stringify(draft.value) : ''
    connectionMessage.value = ''
  },
  { immediate: true },
)

const headersRecord = () => {
  const headers: Record<string, string> = {}
  for (const header of draft.value?.headers || []) {
    const key = header.key.trim()
    if (!key || forbiddenHeaders.has(key.toLowerCase())) continue
    headers[key] = header.value
  }
  return headers
}

const validateDraft = () => {
  if (!draft.value?.name.trim()) return t('providerNameRequired')
  if (!draft.value.baseURL.trim()) return t('baseURLRequired')
  try {
    const url = new URL(draft.value.baseURL)
    if (!['http:', 'https:'].includes(url.protocol)) return t('baseURLInvalid')
  } catch {
    return t('baseURLInvalid')
  }
  if (!draft.value.apiKey.trim()) return t('modelFetchApiKeyRequired')
  return ''
}

const connectionError = (error: unknown) => {
  if (error instanceof ModelCatalogError && error.status) {
    return `HTTP ${error.status}: ${error.message}`
  }
  if (error instanceof DOMException && error.name === 'AbortError') return t('modelFetchTimeout')
  if (error instanceof TypeError) return t('connectionCorsOrNetworkError')
  return error instanceof Error ? error.message : t('modelFetchFailed')
}

const runCatalogRequest = async () => {
  const validationError = validateDraft()
  if (validationError) throw new Error(validationError)
  const controller = new AbortController()
  const timeout = window.setTimeout(() => controller.abort(), draft.value!.timeoutMs || 60000)
  try {
    return await fetchModelCatalog({
      provider: 'official',
      apiKey: draft.value!.apiKey,
      baseURL: draft.value!.baseURL,
      headers: headersRecord(),
      signal: controller.signal,
    })
  } finally {
    window.clearTimeout(timeout)
  }
}

const selectProfile = (id: string) => {
  if (dirty.value && !window.confirm(t('discardUnsavedChanges'))) return
  setActiveProfile(id)
}

const addProvider = () => {
  if (dirty.value && !window.confirm(t('discardUnsavedChanges'))) return
  addProfile(selectedTemplate.value)
}

const deleteProvider = () => {
  if (!draft.value || !window.confirm(t('confirmDeleteProvider'))) return
  removeProfile(draft.value.id)
}

const saveProfile = () => {
  if (!draft.value) return
  const error = validateDraft()
  if (error) {
    connectionOk.value = false
    connectionMessage.value = error
    return
  }
  if (!enabledModels.value.some(model => model.id === draft.value!.defaultModel)) {
    draft.value.defaultModel = enabledModels.value[0]?.id || ''
  }
  updateProfile(draft.value)
  savedSnapshot.value = JSON.stringify(draft.value)
  connectionMessage.value = t('providerSaved')
  connectionOk.value = true
}

const resetDraft = () => {
  draft.value = activeProfile.value ? structuredClone(activeProfile.value) : null
  savedSnapshot.value = draft.value ? JSON.stringify(draft.value) : ''
  connectionMessage.value = ''
}

const testConnection = async () => {
  testing.value = true
  connectionMessage.value = ''
  const startedAt = performance.now()
  try {
    const models = await runCatalogRequest()
    connectionOk.value = true
    connectionMessage.value = t('connectionSuccessDetail', {
      duration: Math.round(performance.now() - startedAt),
      count: models.length,
    })
    if (draft.value) {
      draft.value.lastCheckedAt = new Date().toISOString()
      draft.value.lastConnectionStatus = 'success'
    }
  } catch (error) {
    connectionOk.value = false
    connectionMessage.value = connectionError(error)
    if (draft.value) {
      draft.value.lastCheckedAt = new Date().toISOString()
      draft.value.lastConnectionStatus = 'failed'
    }
  } finally {
    testing.value = false
  }
}

const fetchModels = async () => {
  fetching.value = true
  connectionMessage.value = ''
  try {
    catalogModels.value = await runCatalogRequest()
    pickerOpen.value = true
    connectionOk.value = true
    if (draft.value) {
      draft.value.lastCheckedAt = new Date().toISOString()
      draft.value.lastConnectionStatus = 'success'
    }
    connectionMessage.value = t('modelsFound', { count: catalogModels.value.length })
  } catch (error) {
    connectionOk.value = false
    if (draft.value) {
      draft.value.lastCheckedAt = new Date().toISOString()
      draft.value.lastConnectionStatus = 'failed'
    }
    connectionMessage.value = connectionError(error)
  } finally {
    fetching.value = false
  }
}

const addFetchedModels = (models: string[]) => {
  if (!draft.value) return
  const existing = new Set(draft.value.models.map(model => model.id))
  draft.value.models.push(
    ...models
      .filter(id => !existing.has(id))
      .map(id => ({ ...createProviderModel(id), lastVerifiedAt: new Date().toISOString() })),
  )
  if (!draft.value.defaultModel && draft.value.models.length) draft.value.defaultModel = draft.value.models[0].id
  pickerOpen.value = false
}

const addManualModel = () => {
  const id = manualModel.value.trim()
  if (!draft.value || !id || draft.value.models.some(model => model.id === id)) return
  draft.value.models.push(createProviderModel(id))
  if (!draft.value.defaultModel) draft.value.defaultModel = id
  manualModel.value = ''
}

const removeModel = (id: string) => {
  if (!draft.value) return
  draft.value.models = draft.value.models.filter(model => model.id !== id)
  if (draft.value.defaultModel === id)
    draft.value.defaultModel = draft.value.models.find(model => model.enabled)?.id || ''
}

const setDefaultModel = (id: string) => {
  if (draft.value) draft.value.defaultModel = id
}

const addHeader = () => draft.value?.headers.push({ id: crypto.randomUUID(), key: '', value: '' })
const removeHeader = (id: string) => {
  if (draft.value) draft.value.headers = draft.value.headers.filter(header => header.id !== id)
}

const capabilityClass = (model: ProviderModel, capability: ModelCapability) => {
  const state = getResolvedCapability(model, capability)
  if (state === 'yes') return 'bg-success/15 text-success'
  if (state === 'no') return 'bg-bg-tertiary text-secondary opacity-50'
  return 'bg-warning/15 text-warning'
}

const capabilityTitle = (model: ProviderModel, capability: ModelCapability, label: string) => {
  const state = getResolvedCapability(model, capability)
  const automatic = model.capabilityOverrides?.[capability] === undefined ? t('automatic') : t('manualOverride')
  return `${t(label)}: ${t(`capability_${state}`)} · ${automatic} · ${t('clickToCycle')}`
}
</script>
