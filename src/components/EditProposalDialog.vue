<template>
  <Teleport to="body">
    <div
      v-if="proposal"
      class="fixed inset-0 z-50 flex items-center justify-center bg-black/45 p-3 backdrop-blur-[2px]"
      role="presentation"
      @click.self="emit('cancel')"
    >
      <section
        class="flex max-h-[86vh] w-full max-w-lg flex-col overflow-hidden rounded-xl border border-border bg-bg-secondary shadow-2xl"
        role="dialog"
        aria-modal="true"
        :aria-label="title"
      >
        <header class="flex items-center justify-between gap-3 border-b border-border p-4">
          <div class="min-w-0">
            <h3 class="text-base font-semibold text-main">{{ title }}</h3>
            <p class="mt-1 text-xs text-secondary">{{ subtitle }}</p>
          </div>
          <button
            class="rounded-md p-1.5 text-secondary hover:bg-surface"
            :aria-label="cancelLabel"
            @click="emit('cancel')"
          >
            <X :size="18" />
          </button>
        </header>

        <div class="flex items-center justify-between gap-2 border-b border-border px-4 py-2 text-xs">
          <span :class="riskClass">{{ riskLabel }}</span>
          <span class="text-secondary">{{ changeSummary }}</span>
        </div>

        <div class="min-h-32 flex-1 overflow-y-auto p-3">
          <div
            class="rounded-md border border-border bg-surface p-2 font-mono text-xs leading-relaxed whitespace-pre-wrap"
          >
            <div v-for="(line, index) in diff" :key="index" class="rounded-sm px-2" :class="lineClass(line.type)">
              <span class="mr-1 opacity-60 select-none">{{
                line.type === 'added' ? '+' : line.type === 'removed' ? '-' : ' '
              }}</span
              >{{ line.text }}
            </div>
          </div>
          <details class="mt-3 text-xs text-secondary">
            <summary class="cursor-pointer">{{ fullTextLabel }}</summary>
            <div class="mt-2 grid gap-2">
              <div>
                <div class="mb-1 font-semibold">{{ beforeLabel }}</div>
                <pre class="m-0 max-h-32 overflow-auto rounded-md bg-surface p-2 whitespace-pre-wrap">{{
                  proposal.beforeText
                }}</pre>
              </div>
              <div>
                <div class="mb-1 font-semibold">{{ afterLabel }}</div>
                <pre class="m-0 max-h-32 overflow-auto rounded-md bg-surface p-2 whitespace-pre-wrap">{{
                  proposal.afterText
                }}</pre>
              </div>
            </div>
          </details>
        </div>

        <footer class="flex flex-wrap items-center justify-between gap-2 border-t border-border p-3">
          <button class="rounded-md px-2 py-2 text-xs text-secondary hover:bg-surface" @click="copyAfterText">
            {{ copyLabel }}
          </button>
          <div class="flex gap-2">
            <button
              class="rounded-md border border-border px-3 py-2 text-sm text-main hover:bg-surface"
              @click="emit('cancel')"
            >
              {{ cancelLabel }}
            </button>
            <button
              class="rounded-md bg-accent px-3 py-2 text-sm font-medium text-white hover:bg-accent-hover"
              @click="emit('accept')"
            >
              {{ acceptLabel }}
            </button>
          </div>
        </footer>
      </section>
    </div>
  </Teleport>
</template>

<script setup lang="ts">
import { X } from 'lucide-vue-next'
import { computed } from 'vue'

import { buildTextDiff, countDiffChanges, DiffLine, TextChangeProposal } from '@/utils/textProposal'

const props = defineProps<{
  proposal: TextChangeProposal | null
  title: string
  subtitle: string
  beforeLabel: string
  afterLabel: string
  fullTextLabel: string
  copyLabel: string
  cancelLabel: string
  acceptLabel: string
  riskLabels: Record<TextChangeProposal['risk'], string>
  changeSummaryTemplate: string
}>()

const emit = defineEmits<{
  accept: []
  cancel: []
  copied: []
}>()

const diff = computed(() => (props.proposal ? buildTextDiff(props.proposal.beforeText, props.proposal.afterText) : []))
const changes = computed(() => countDiffChanges(diff.value))
const riskLabel = computed(() => (props.proposal ? props.riskLabels[props.proposal.risk] : ''))
const riskClass = computed(() => {
  if (!props.proposal) return ''
  if (props.proposal.risk === 'high') return 'text-danger'
  if (props.proposal.risk === 'medium') return 'text-warning'
  return 'text-success'
})
const changeSummary = computed(() =>
  props.changeSummaryTemplate
    .replace('{added}', String(changes.value.added))
    .replace('{removed}', String(changes.value.removed)),
)

const lineClass = (type: DiffLine['type']) => {
  if (type === 'added') return 'bg-success/15 text-success'
  if (type === 'removed') return 'bg-danger/15 text-danger'
  return 'text-secondary'
}

const copyAfterText = async () => {
  if (!props.proposal) return
  await navigator.clipboard.writeText(props.proposal.afterText)
  emit('copied')
}
</script>
