<template>
  <Teleport to="body">
    <div
      v-if="open"
      class="fixed inset-0 z-50 flex items-center justify-center bg-black/45 p-4 backdrop-blur-[2px]"
      role="presentation"
      @click.self="emit('close')"
    >
      <section
        class="flex max-h-[80vh] w-full max-w-lg flex-col overflow-hidden rounded-xl border border-border bg-bg-secondary shadow-2xl"
        role="dialog"
        aria-modal="true"
        :aria-label="title"
      >
        <header class="flex items-center justify-between gap-3 border-b border-border p-4">
          <div class="min-w-0">
            <h3 class="truncate text-base font-semibold text-main">{{ title }}</h3>
            <p class="mt-0.5 text-xs text-secondary">{{ subtitle }}</p>
          </div>
          <button
            class="rounded-md p-1.5 text-secondary transition-colors hover:bg-surface hover:text-main"
            :aria-label="closeLabel"
            @click="emit('close')"
          >
            <X :size="18" />
          </button>
        </header>

        <div class="border-b border-border p-3">
          <label class="flex items-center gap-2 rounded-md border border-border bg-surface px-3 py-2">
            <Search :size="15" class="shrink-0 text-secondary" />
            <input
              v-model="query"
              class="min-w-0 flex-1 bg-transparent text-sm text-main outline-none placeholder:text-secondary"
              :placeholder="searchPlaceholder"
              autofocus
            />
          </label>
          <div class="mt-2 flex items-center justify-between text-xs text-secondary">
            <span>{{ resultSummary }}</span>
            <button class="text-accent hover:underline" @click="toggleAllVisible">
              {{ allVisibleSelected ? clearVisibleLabel : selectVisibleLabel }}
            </button>
          </div>
        </div>

        <div class="min-h-40 flex-1 overflow-y-auto p-2">
          <label
            v-for="model in filteredModels"
            :key="model"
            class="flex cursor-pointer items-center gap-3 rounded-md px-3 py-2.5 hover:bg-surface"
          >
            <input v-model="selected" type="checkbox" :value="model" class="h-4 w-4 accent-accent" />
            <span class="min-w-0 flex-1 text-sm break-all text-main">{{ model }}</span>
            <span v-if="existingModels.includes(model)" class="shrink-0 text-xs text-secondary">{{ addedLabel }}</span>
          </label>
          <div
            v-if="filteredModels.length === 0"
            class="flex min-h-40 items-center justify-center text-sm text-secondary"
          >
            {{ emptyLabel }}
          </div>
        </div>

        <footer class="flex items-center justify-between gap-3 border-t border-border p-3">
          <span class="text-xs text-secondary">{{ selected.length }} {{ selectedLabel }}</span>
          <div class="flex gap-2">
            <button
              class="rounded-md border border-border px-3 py-2 text-sm text-main hover:bg-surface"
              @click="emit('close')"
            >
              {{ cancelLabel }}
            </button>
            <button
              :disabled="selected.length === 0"
              class="rounded-md bg-accent px-3 py-2 text-sm font-medium text-white hover:bg-accent-hover disabled:cursor-not-allowed disabled:opacity-50"
              @click="confirm"
            >
              {{ addSelectedLabel }}
            </button>
          </div>
        </footer>
      </section>
    </div>
  </Teleport>
</template>

<script setup lang="ts">
import { Search, X } from 'lucide-vue-next'
import { computed, ref, watch } from 'vue'

const props = defineProps<{
  open: boolean
  models: string[]
  existingModels: string[]
  title: string
  subtitle: string
  searchPlaceholder: string
  closeLabel: string
  selectVisibleLabel: string
  clearVisibleLabel: string
  addedLabel: string
  emptyLabel: string
  selectedLabel: string
  cancelLabel: string
  addSelectedLabel: string
  resultSummary: string
}>()

const emit = defineEmits<{
  close: []
  confirm: [models: string[]]
}>()

const query = ref('')
const selected = ref<string[]>([])

const filteredModels = computed(() => {
  const keyword = query.value.trim().toLowerCase()
  if (!keyword) return props.models
  return props.models.filter(model => model.toLowerCase().includes(keyword))
})

const allVisibleSelected = computed(
  () => filteredModels.value.length > 0 && filteredModels.value.every(model => selected.value.includes(model)),
)

watch(
  () => props.open,
  open => {
    if (!open) return
    query.value = ''
    selected.value = props.models.filter(model => !props.existingModels.includes(model))
  },
)

const toggleAllVisible = () => {
  if (allVisibleSelected.value) {
    const visible = new Set(filteredModels.value)
    selected.value = selected.value.filter(model => !visible.has(model))
    return
  }
  selected.value = [...new Set([...selected.value, ...filteredModels.value])]
}

const confirm = () => emit('confirm', selected.value)
</script>
