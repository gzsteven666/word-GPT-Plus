<template>
  <Teleport to="body">
    <div
      v-if="open"
      class="fixed inset-0 z-50 flex items-center justify-center bg-black/45 p-3 backdrop-blur-[2px]"
      role="presentation"
      @click.self="emit('cancel')"
    >
      <section
        class="w-full max-w-md overflow-hidden rounded-xl border border-border bg-bg-secondary shadow-2xl"
        role="dialog"
        aria-modal="true"
        :aria-label="title"
      >
        <header class="flex items-center justify-between gap-3 border-b border-border p-4">
          <h3 class="text-base font-semibold text-main">{{ title }}</h3>
          <button
            class="rounded-md p-1.5 text-secondary hover:bg-surface"
            :aria-label="cancelLabel"
            @click="emit('cancel')"
          >
            <X :size="18" />
          </button>
        </header>

        <p class="p-4 text-sm leading-relaxed whitespace-pre-wrap text-main">{{ message }}</p>

        <footer class="flex justify-end gap-2 border-t border-border p-3">
          <button
            class="rounded-md border border-border px-3 py-2 text-sm text-main hover:bg-surface"
            @click="emit('cancel')"
          >
            {{ cancelLabel }}
          </button>
          <button
            class="rounded-md bg-accent px-3 py-2 text-sm font-medium text-white hover:bg-accent-hover"
            @click="emit('confirm')"
          >
            {{ confirmLabel }}
          </button>
        </footer>
      </section>
    </div>
  </Teleport>
</template>

<script setup lang="ts">
import { X } from 'lucide-vue-next'

defineProps<{
  open: boolean
  title: string
  message: string
  confirmLabel: string
  cancelLabel: string
}>()

const emit = defineEmits<{
  confirm: []
  cancel: []
}>()
</script>
