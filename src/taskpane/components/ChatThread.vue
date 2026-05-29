<script setup lang="ts">
import { nextTick, ref, watch } from 'vue'
import { useI18n } from 'vue-i18n'
import ActionButton from '@/taskpane/components/ActionButton.vue'

export interface ChatMessage {
  role: 'user' | 'ai'
  text: string
}

interface Props {
  messages: ChatMessage[]
  loading?: boolean
  placeholder?: string
  emptyHint?: string
}

const props = withDefaults(defineProps<Props>(), {
  loading: false,
  placeholder: '',
  emptyHint: '',
})

const emit = defineEmits<{ (e: 'send', text: string): void }>()

const { t } = useI18n()
const draft = ref('')
const scroller = ref<HTMLElement | null>(null)

function submit(): void {
  const text = draft.value.trim()
  if (!text || props.loading) return
  emit('send', text)
  draft.value = ''
}

watch(
  () => props.messages.length,
  async () => {
    await nextTick()
    if (scroller.value) scroller.value.scrollTop = scroller.value.scrollHeight
  },
)
</script>

<template>
  <div class="chat">
    <div ref="scroller" class="chat__scroll">
      <p v-if="messages.length === 0 && emptyHint" class="syn-muted chat__empty">
        {{ emptyHint }}
      </p>
      <div v-for="(m, i) in messages" :key="i" :class="['chat__bubble', `chat__bubble--${m.role}`]">
        {{ m.text }}
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
.chat__scroll {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-2);
  max-height: 320px;
  overflow-y: auto;
  padding: var(--syn-space-1);
}
.chat__empty {
  margin: 0;
  text-align: center;
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
.chat__dots {
  margin-right: var(--syn-space-1);
}
.chat__composer {
  display: flex;
  gap: var(--syn-space-2);
  align-items: flex-end;
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
