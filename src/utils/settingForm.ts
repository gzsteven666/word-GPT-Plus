import { Ref, ref } from 'vue'

import { localStorageKey } from './enum'
import { Setting_Names, SettingNames, settingPreset } from './settingPreset'

type SettingForm = {
  [K in SettingNames]: (typeof settingPreset)[K]['defaultValue']
}

type SettingValue = string | number | string[]

function initializeSettings(): Record<string, SettingValue> {
  const settings: Record<string, SettingValue> = {}

  for (const key of Setting_Names) {
    const preset = settingPreset[key]

    if (preset.getFunc) {
      settings[key] = preset.getFunc()
    } else {
      const storageKey = preset.saveKey || key
      const storedValue = localStorage.getItem(storageKey)
      settings[key] = storedValue ?? preset.defaultValue
    }
  }

  // Special case for legacy support
  if (settings.api === 'palm') {
    settings.api = 'gemini'
    localStorage.setItem(localStorageKey.api, 'gemini')
  }

  // Migrate untouched legacy output limits to automatic mode. A user-customized
  // value is preserved; 0 lets the runtime choose a task-appropriate budget.
  if (!localStorage.getItem('automaticTokenDefaultsV2')) {
    const legacyDefaults: Partial<Record<SettingNames, number>> = {
      officialMaxTokens: 800,
      azureMaxTokens: 800,
      geminiMaxTokens: 800,
      groqMaxTokens: 1024,
    }
    for (const [key, legacyValue] of Object.entries(legacyDefaults) as [SettingNames, number][]) {
      if (Number(settings[key]) === legacyValue) {
        settings[key] = 0
        localStorage.setItem(settingPreset[key].saveKey || key, '0')
      }
    }
    localStorage.setItem('automaticTokenDefaultsV2', '1')
  }

  return settings
}

function useSettingForm() {
  return ref(initializeSettings()) as Ref<SettingForm>
}

export default useSettingForm
