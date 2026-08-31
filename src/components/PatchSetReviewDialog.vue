<template>
  <div v-if="patchSet" class="fixed inset-0 z-50 flex items-center justify-center bg-black/30 p-3">
    <div
      class="flex max-h-[90%] w-full max-w-xl flex-col gap-2 overflow-hidden rounded-md border border-border bg-surface p-3 shadow-xl"
    >
      <div class="flex items-center justify-between">
        <h2 class="m-0 text-sm font-semibold text-main">Review document changes</h2>
        <span class="text-xs text-secondary">{{ selectedIds.size }}/{{ patchSet.operations.length }} selected</span>
      </div>
      <div class="flex flex-col gap-2 overflow-y-auto">
        <label
          v-for="operation in patchSet.operations"
          :key="operation.id"
          class="flex cursor-pointer flex-col gap-1 rounded-md border border-border-secondary p-2 text-xs"
        >
          <div class="flex items-center gap-2">
            <input
              type="checkbox"
              :checked="selectedIds.has(operation.id)"
              :disabled="operation.status === 'conflicted'"
              @change="toggle(operation.id)"
            />
            <span class="font-semibold text-main">{{ operation.type }}</span>
            <span class="text-secondary">{{ operation.risk }} risk</span>
          </div>
          <div class="grid grid-cols-2 gap-2">
            <pre class="m-0 overflow-x-auto rounded bg-danger/10 p-1 whitespace-pre-wrap text-danger">{{
              operation.beforeText
            }}</pre>
            <pre class="m-0 overflow-x-auto rounded bg-success/10 p-1 whitespace-pre-wrap text-success">{{
              operation.replacementText
            }}</pre>
          </div>
          <div class="flex flex-wrap gap-1 text-secondary">
            <span
              v-for="(line, index) in buildTextDiff(operation.beforeText, operation.replacementText)"
              :key="index"
              :class="line.type === 'added' ? 'text-success' : line.type === 'removed' ? 'text-danger' : ''"
            >
              {{ line.text }}
            </span>
          </div>
        </label>
      </div>
      <div class="flex justify-end gap-2">
        <button class="rounded border border-border px-3 py-1 text-xs text-secondary" @click="$emit('cancel')">
          Cancel
        </button>
        <button
          class="rounded bg-accent px-3 py-1 text-xs text-white disabled:opacity-50"
          :disabled="selectedIds.size === 0"
          @click="$emit('apply', [...selectedIds])"
        >
          Apply selected
        </button>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, watch } from 'vue'

import type { DocumentPatchSet } from '@/utils/documentPatch'
import { buildTextDiff } from '@/utils/textProposal'

const props = defineProps<{ patchSet: DocumentPatchSet | null }>()
defineEmits<{
  apply: [operationIds: string[]]
  cancel: []
}>()

const selectedIds = ref<Set<string>>(new Set())
watch(
  () => props.patchSet,
  patchSet => {
    selectedIds.value = new Set(patchSet?.operations.map(operation => operation.id) || [])
  },
  { immediate: true },
)
const toggle = (operationId: string) => {
  const next = new Set(selectedIds.value)
  if (next.has(operationId)) next.delete(operationId)
  else next.add(operationId)
  selectedIds.value = next
}
</script>
