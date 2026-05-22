<script setup lang="ts">
import { computed, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import ActionButton from '@/taskpane/components/ActionButton.vue'
import LanguagePicker from '@/taskpane/components/LanguagePicker.vue'
import TonePicker from '@/taskpane/components/TonePicker.vue'
import Toast from '@/taskpane/components/Toast.vue'
import { useOutlookItem } from '@/taskpane/composables/useOutlookItem'
import { useSynaplanClient } from '@/taskpane/composables/useSynaplanClient'
import { go } from '@/taskpane/router'

const { t } = useI18n()
const { item } = useOutlookItem()
const { call } = useSynaplanClient()

type ActionKey = 'summarise' | 'translate' | 'reply' | 'classify' | 'save' | null
const active = ref<ActionKey>(null)
const result = ref<string>('')
const error = ref<string | null>(null)
const targetLang = ref<'auto' | 'en' | 'de' | 'fr' | 'es' | 'it' | 'zh' | 'ar'>('en')
const tone = ref<'formal' | 'concise' | 'friendly'>('concise')
const question = ref('')
const askHistory = ref<{ q: string; a: string }[]>([])

const senderEmail = computed(() => item.value.from ?? '')

async function run<T>(key: NonNullable<ActionKey>, fn: () => Promise<T | null>): Promise<T | null> {
  active.value = key
  error.value = null
  try {
    return await fn()
  } catch (err) {
    error.value = err instanceof Error ? err.message : String(err)
    return null
  } finally {
    active.value = null
  }
}

async function summarise(): Promise<void> {
  const r = await run('summarise', () =>
    call((c) =>
      c.summarise({
        subject: item.value.subject,
        body: item.value.bodyText,
        from: item.value.from,
        to: item.value.to,
      }),
    ),
  )
  if (r) result.value = r.bullets.map((b) => `• ${b}`).join('\n')
}

async function translate(): Promise<void> {
  const r = await run('translate', () =>
    call((c) =>
      c.translate({
        text: item.value.bodyText,
        targetLanguage: targetLang.value === 'auto' ? 'en' : targetLang.value,
      }),
    ),
  )
  if (r) result.value = r.translation
}

async function draftReply(): Promise<void> {
  const r = await run('reply', () =>
    call((c) =>
      c.draftReply({
        subject: item.value.subject,
        body: item.value.bodyText,
        tone: tone.value,
        language: targetLang.value === 'auto' ? 'en' : targetLang.value,
      }),
    ),
  )
  if (r) {
    try {
      Office.context.mailbox.item?.displayReplyForm({ htmlBody: r.htmlBody })
    } catch (err) {
      error.value = err instanceof Error ? err.message : String(err)
    }
  }
}

async function classify(): Promise<void> {
  const r = await run('classify', () =>
    call((c) =>
      c.classify({ subject: item.value.subject, body: item.value.bodyText, from: item.value.from }),
    ),
  )
  if (r) result.value = `${r.category} (${Math.round(r.confidence * 100)}%) — ${r.reasoning}`
}

async function saveToRag(): Promise<void> {
  await run('save', async () => {
    const r = await call((c) =>
      c.fileUpload({
        filename: item.value.subject + '.eml',
        contentBase64: btoa(unescape(encodeURIComponent(item.value.bodyText))),
        mimeType: 'message/rfc822',
        groupId: senderEmail.value ? `contact:${senderEmail.value.toLowerCase()}` : undefined,
        metadata: { from: item.value.from ?? '', subject: item.value.subject },
        // Synaplan's upload now extracts as part of the same request — no
        // separate fileProcess call needed (and the interface no longer
        // exposes one). Use 'vectorize' if the user opts into RAG retrieval.
        processLevel: 'extract',
      }),
    )
    if (r) {
      result.value = `Saved file #${r.fileId}.`
    }
    return r
  })
}

async function ask(): Promise<void> {
  const q = question.value.trim()
  if (!q) return
  const conversationId = item.value.conversationId ?? `synamail:${item.value.subject}`
  const r = await run('classify', () =>
    call((c) => c.ask({ conversationId, question: q, emailContext: item.value.bodyText })),
  )
  if (r) {
    askHistory.value.push({ q, a: r.answer })
    question.value = ''
  }
}
</script>

<template>
  <section class="read">
    <header class="read__header">
      <h2 class="read__subject">
        {{ item.subject || '—' }}
      </h2>
      <p class="syn-muted">
        <span v-if="item.from">From: {{ item.from }}</span>
      </p>
      <button v-if="senderEmail" class="read__contact" type="button" @click="go('contact-kb')">
        {{ t('read.contactPill', { email: senderEmail }) }}
      </button>
    </header>

    <div class="syn-stack">
      <ActionButton :loading="active === 'summarise'" @click="summarise">
        {{ t('read.actions.summarise') }}
      </ActionButton>

      <div class="syn-row">
        <ActionButton :loading="active === 'translate'" :block="false" @click="translate">
          {{ t('read.actions.translate') }}
        </ActionButton>
        <LanguagePicker v-model="targetLang" />
      </div>

      <div class="syn-row">
        <ActionButton :loading="active === 'reply'" :block="false" @click="draftReply">
          {{ t('read.actions.draftReply') }}
        </ActionButton>
        <TonePicker v-model="tone" />
      </div>

      <ActionButton :loading="active === 'classify'" @click="classify">
        {{ t('read.actions.classify') }}
      </ActionButton>

      <ActionButton :loading="active === 'save'" @click="saveToRag">
        {{ t('read.actions.saveToRag') }}
      </ActionButton>
    </div>

    <pre v-if="result" class="read__result">{{ result }}</pre>
    <Toast v-if="error" kind="error" :message="error" />

    <section class="read__ask">
      <div v-for="(turn, i) in askHistory" :key="i" class="read__turn">
        <p><strong>You:</strong> {{ turn.q }}</p>
        <p><strong>Synaplan:</strong> {{ turn.a }}</p>
      </div>
      <div class="syn-row">
        <input
          v-model="question"
          type="text"
          :placeholder="t('read.askPlaceholder')"
          class="read__ask-input"
          @keyup.enter="ask"
        />
        <ActionButton :block="false" :loading="active === 'classify'" @click="ask">
          {{ t('read.actions.ask') }}
        </ActionButton>
      </div>
    </section>
  </section>
</template>

<style scoped>
.read {
  padding: var(--syn-space-3);
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-3);
}
.read__header {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-1);
}
.read__subject {
  margin: 0;
  font-size: var(--syn-font-size-lg);
}
.read__contact {
  align-self: flex-start;
  border: 1px solid var(--syn-border);
  background: var(--syn-surface);
  color: var(--syn-text);
  padding: var(--syn-space-1) var(--syn-space-2);
  border-radius: 999px;
  font-size: var(--syn-font-size-sm);
}
.read__result {
  background: var(--syn-surface);
  padding: var(--syn-space-3);
  border-radius: var(--syn-radius-md);
  white-space: pre-wrap;
  margin: 0;
  font-family: inherit;
  font-size: var(--syn-font-size-sm);
}
.read__ask {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-2);
}
.read__turn {
  padding: var(--syn-space-2);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-sm);
}
.read__ask-input {
  flex: 1;
  padding: var(--syn-space-2);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-sm);
  font-family: inherit;
}
</style>
