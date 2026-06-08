<script setup lang="ts">
import { computed, nextTick, ref, watch } from 'vue'
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

// True when the last message is an AI bubble still waiting for its first token
// (the caller streams into it). The thinking indicator then lives in that
// bubble instead of a separate one, so they never show at the same time.
const pendingAi = computed(() => {
  const last = props.messages[props.messages.length - 1]
  return !!last && last.role === 'ai' && !last.text
})

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
        <template v-if="m.role === 'ai'">
          <!-- While a streamed bubble is still empty, it shows the thinking
               indicator; once tokens arrive it renders Markdown. renderMarkdown
               sanitises the text, so v-html is XSS-safe here. -->
          <span v-if="!m.text && loading && i === messages.length - 1" class="chat__loading">
            <span class="chat__dots" aria-hidden="true">…</span>
            {{ t('home.chat.thinking') }}
          </span>
          <!-- eslint-disable-next-line vue/no-v-html -->
          <div v-else class="syn-md" v-html="renderMarkdown(m.text)" />
        </template>
        <template v-else>{{ m.text }}</template>
      </div>
      <div v-if="loading && !pendingAi" class="chat__bubble chat__bubble--ai chat__bubble--loading">
        <span class="chat__dots" aria-hidden="true">…</span>
        {{ t('home.chat.thinking') }}
      </div>
    </div>

    <div class="chat__composer" :class="{ 'chat__composer--bordered': messages.length }">
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
/* Markdown formatting comes from the shared `.syn-md` rules in app.css. */
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
  background: var(--syn-surface);
}
/* Only divide the composer from the thread once messages exist — on the empty
   initial state the line under the non-foldable "Chat" header looks out of
   place next to the foldable accordions below. */
.chat__composer--bordered {
  padding-top: var(--syn-space-2);
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
