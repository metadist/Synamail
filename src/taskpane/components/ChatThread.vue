<script setup lang="ts">
import { nextTick, ref, watch } from 'vue'
import { useI18n } from 'vue-i18n'
import ActionButton from '@/taskpane/components/ActionButton.vue'
import { renderMarkdown } from '@shared/markdown'

export interface ChatMessage {
  role: 'user' | 'ai'
  text: string
}

interface Props {
  messages: ChatMessage[]
  loading?: boolean
  placeholder?: string
  emptyHint?: string
  /** Seed the composer so Send is active and the user sees a concrete example. */
  initialDraft?: string
}

const props = withDefaults(defineProps<Props>(), {
  loading: false,
  placeholder: '',
  emptyHint: '',
  initialDraft: '',
})

const emit = defineEmits<{ (e: 'send', text: string): void; (e: 'reset'): void }>()

const { t } = useI18n()
const draft = ref(props.initialDraft)
const root = ref<HTMLElement | null>(null)

// The composer is seeded with an example so Send is active and users see they
// can type here. The first time they focus the still-untouched sample, clear it
// so they don't have to delete it by hand. Once they've focused or edited it,
// we never auto-clear again — so it's never annoying on later interactions.
const sampleActive = ref(Boolean(props.initialDraft))

function onFocus(): void {
  if (sampleActive.value && draft.value === props.initialDraft) {
    draft.value = ''
  }
  sampleActive.value = false
}

function onInput(): void {
  sampleActive.value = false
}

function submit(): void {
  const text = draft.value.trim()
  if (!text || props.loading) return
  emit('send', text)
  draft.value = ''
  sampleActive.value = false
}

function scrollToTop(): void {
  root.value?.scrollIntoView({ behavior: 'smooth', block: 'start' })
}

// Keep the newest reply — and the sticky composer — in view as the thread grows.
function scrollToLatest(): void {
  try {
    globalThis.scrollTo?.({ top: document.documentElement.scrollHeight, behavior: 'smooth' })
  } catch {
    /* jsdom / non-browser env has no real scrolling */
  }
}

watch(
  () => props.messages.length,
  async () => {
    await nextTick()
    scrollToLatest()
  },
)
</script>

<template>
  <div ref="root" class="chat">
    <div v-if="messages.length" class="chat__toolbar">
      <button type="button" class="chat__tool" @click="scrollToTop">
        ↑ {{ t('home.chat.scrollTop') }}
      </button>
      <button type="button" class="chat__tool" @click="$emit('reset')">
        {{ t('home.chat.reset') }}
      </button>
    </div>

    <div class="chat__scroll">
      <p v-if="messages.length === 0 && emptyHint" class="syn-muted chat__empty">
        {{ emptyHint }}
      </p>
      <div v-for="(m, i) in messages" :key="i" :class="['chat__bubble', `chat__bubble--${m.role}`]">
        <!-- AI replies are Markdown; renderMarkdown escapes first, so v-html is
             XSS-safe here. User messages stay plain text. -->
        <!-- eslint-disable-next-line vue/no-v-html -->
        <div v-if="m.role === 'ai'" class="chat__md" v-html="renderMarkdown(m.text)" />
        <template v-else>{{ m.text }}</template>
      </div>
      <div v-if="loading" class="chat__bubble chat__bubble--ai chat__bubble--loading">
        <span class="chat__dots" aria-hidden="true">…</span>
        {{ t('home.chat.thinking') }}
      </div>
    </div>

    <div class="chat__composer">
      <textarea
        v-model="draft"
        rows="2"
        class="chat__input"
        :placeholder="placeholder || t('home.chat.placeholder')"
        @focus="onFocus"
        @input="onInput"
        @keydown.enter.exact.prevent="submit"
      />
      <ActionButton
        variant="primary"
        :block="false"
        :loading="loading"
        :disabled="!draft.trim()"
        @click="submit"
      >
        {{ t('home.chat.send') }}
      </ActionButton>
    </div>
  </div>
</template>

<style scoped>
.chat {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-2);
}
.chat__toolbar {
  position: sticky;
  top: 0;
  z-index: 2;
  display: flex;
  justify-content: flex-end;
  gap: var(--syn-space-3);
  padding-bottom: var(--syn-space-1);
  background: var(--syn-surface);
  border-bottom: 1px solid var(--syn-border);
}
.chat__tool {
  background: none;
  border: 0;
  padding: 0;
  color: var(--syn-brand-600);
  font-size: var(--syn-font-size-sm);
  font-weight: 500;
}
.chat__tool:hover {
  text-decoration: underline;
}
.chat__scroll {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-2);
  padding: var(--syn-space-1);
}
.chat__empty {
  margin: 0;
  text-align: left;
  font-size: var(--syn-font-size-sm);
}
.chat__bubble {
  padding: var(--syn-space-2) var(--syn-space-3);
  border-radius: var(--syn-radius-md);
  font-size: var(--syn-font-size-sm);
  white-space: pre-wrap;
  word-break: break-word;
  max-width: 92%;
}
.chat__bubble--user {
  align-self: flex-end;
  background: var(--syn-brand-600);
  color: white;
}
.chat__bubble--ai {
  align-self: flex-start;
  background: var(--syn-surface);
  border: 1px solid var(--syn-border);
}
.chat__bubble--loading {
  opacity: 0.8;
}
.chat__md :first-child {
  margin-top: 0;
}
.chat__md :last-child {
  margin-bottom: 0;
}
.chat__md p {
  margin: 0 0 var(--syn-space-2);
}
.chat__md ul,
.chat__md ol {
  margin: 0 0 var(--syn-space-2);
  padding-left: 1.25rem;
}
.chat__md .md-h {
  display: block;
  margin: var(--syn-space-2) 0 var(--syn-space-1);
}
.chat__md code {
  background: rgba(127, 127, 127, 0.18);
  padding: 0 4px;
  border-radius: var(--syn-radius-sm);
  font-family: ui-monospace, 'Cascadia Code', 'Consolas', monospace;
  font-size: 0.95em;
}
.chat__md .md-pre {
  background: rgba(127, 127, 127, 0.14);
  padding: var(--syn-space-2);
  border-radius: var(--syn-radius-sm);
  overflow-x: auto;
  margin: 0 0 var(--syn-space-2);
}
.chat__md .md-pre code {
  background: none;
  padding: 0;
}
.chat__md a {
  color: var(--syn-brand-600);
}
.chat__dots {
  margin-right: var(--syn-space-1);
}
.chat__composer {
  position: sticky;
  bottom: 0;
  z-index: 2;
  display: flex;
  gap: var(--syn-space-2);
  align-items: flex-end;
  padding-top: var(--syn-space-2);
  background: var(--syn-surface);
  border-top: 1px solid var(--syn-border);
}
.chat__input {
  flex: 1;
  padding: var(--syn-space-2);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-sm);
  font-family: inherit;
  font-size: var(--syn-font-size-sm);
  resize: vertical;
}
</style>
