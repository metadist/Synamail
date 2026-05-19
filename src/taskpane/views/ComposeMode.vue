<script setup lang="ts">
import { ref } from 'vue'
import { useI18n } from 'vue-i18n'
import ActionButton from '@/taskpane/components/ActionButton.vue'
import LanguagePicker from '@/taskpane/components/LanguagePicker.vue'
import Toast from '@/taskpane/components/Toast.vue'
import { useSynaplanClient } from '@/taskpane/composables/useSynaplanClient'

const { t } = useI18n()
const { call } = useSynaplanClient()
const intent = ref('')
const targetLang = ref<'auto' | 'en' | 'de' | 'fr' | 'es' | 'it' | 'zh' | 'ar'>('en')
const ragQuery = ref('')
const ragHits = ref<{ filename: string; snippet: string; score: number }[]>([])
const active = ref<string | null>(null)
const error = ref<string | null>(null)

function getCompose(): Office.MessageCompose | null {
  const i = Office.context?.mailbox?.item as Office.MessageCompose | undefined
  return i && 'body' in i && typeof (i.body as { setAsync?: unknown }).setAsync === 'function'
    ? i
    : null
}

async function setBody(html: string): Promise<void> {
  const compose = getCompose()
  if (!compose) {
    error.value = 'Compose surface not available'
    return
  }
  await new Promise<void>((resolve, reject) => {
    compose.body.setAsync(html, { coercionType: Office.CoercionType.Html }, (r) =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error),
    )
  })
}

async function setSelected(text: string): Promise<void> {
  const compose = getCompose()
  if (!compose) {
    error.value = 'Compose surface not available'
    return
  }
  await new Promise<void>((resolve, reject) => {
    compose.body.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, (r) =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error),
    )
  })
}

async function getSelected(): Promise<string> {
  const compose = getCompose()
  if (!compose) return ''
  return new Promise<string>((resolve) => {
    compose.body.getAsync(Office.CoercionType.Text, (r) =>
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : ''),
    )
  })
}

async function run<T>(key: string, fn: () => Promise<T>): Promise<T | null> {
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

async function draftFromPrompt(): Promise<void> {
  if (!intent.value.trim()) return
  await run('draft', async () => {
    const r = await call((c) =>
      c.draftReply({
        subject: intent.value,
        body: intent.value,
        tone: 'concise',
        language: targetLang.value === 'auto' ? 'en' : targetLang.value,
      }),
    )
    if (r) await setBody(r.htmlBody)
  })
}

async function transformSelection(kind: 'improve' | 'shorten' | 'translate'): Promise<void> {
  await run(kind, async () => {
    const sel = await getSelected()
    if (!sel.trim()) {
      error.value = 'No text selected'
      return
    }
    if (kind === 'translate') {
      const r = await call((c) =>
        c.translate({
          text: sel,
          targetLanguage: targetLang.value === 'auto' ? 'en' : targetLang.value,
        }),
      )
      if (r) await setSelected(r.translation)
    } else {
      const r = await call((c) => c.summarise({ subject: kind, body: sel }))
      if (r) await setSelected(r.bullets.join(' '))
    }
  })
}

async function searchRag(): Promise<void> {
  if (!ragQuery.value.trim()) return
  await run('rag-search', async () => {
    const r = await call((c) => c.ragSearch({ query: ragQuery.value }))
    if (r)
      ragHits.value = r.map((h) => ({ filename: h.filename, snippet: h.snippet, score: h.score }))
  })
}

async function insertHit(hit: { filename: string; snippet: string }): Promise<void> {
  await setSelected(`\n[${hit.filename}] ${hit.snippet}\n`)
}
</script>

<template>
  <section class="compose">
    <h2>{{ t('compose.title') }}</h2>

    <div class="syn-stack">
      <label for="intent" class="syn-muted">{{ t('compose.intentLabel') }}</label>
      <textarea
        id="intent"
        v-model="intent"
        rows="3"
        :placeholder="t('compose.intentPlaceholder')"
      />
      <div class="syn-row">
        <ActionButton :loading="active === 'draft'" :block="false" @click="draftFromPrompt">
          {{ t('compose.draft') }}
        </ActionButton>
        <LanguagePicker v-model="targetLang" />
      </div>
    </div>

    <div class="syn-row">
      <ActionButton
        :loading="active === 'improve'"
        :block="false"
        @click="transformSelection('improve')"
      >
        {{ t('compose.improve') }}
      </ActionButton>
      <ActionButton
        :loading="active === 'shorten'"
        :block="false"
        @click="transformSelection('shorten')"
      >
        {{ t('compose.shorten') }}
      </ActionButton>
      <ActionButton
        :loading="active === 'translate'"
        :block="false"
        @click="transformSelection('translate')"
      >
        {{ t('compose.translate') }}
      </ActionButton>
    </div>

    <div class="syn-stack">
      <label for="rag" class="syn-muted">{{ t('compose.insertFromRag') }}</label>
      <input
        id="rag"
        v-model="ragQuery"
        type="text"
        :placeholder="t('compose.ragSearchPlaceholder')"
        @keyup.enter="searchRag"
      />
      <ul class="compose__hits">
        <li v-for="(h, i) in ragHits" :key="i">
          <button type="button" @click="insertHit(h)">
            {{ h.filename }} — {{ h.snippet.slice(0, 60) }}…
          </button>
        </li>
      </ul>
    </div>

    <Toast v-if="error" kind="error" :message="error" />
  </section>
</template>

<style scoped>
.compose {
  padding: var(--syn-space-3);
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-4);
}
textarea,
input {
  padding: var(--syn-space-2);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-sm);
  font-family: inherit;
  resize: vertical;
}
.compose__hits {
  list-style: none;
  padding: 0;
  margin: 0;
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-1);
}
.compose__hits button {
  width: 100%;
  text-align: left;
  background: var(--syn-surface);
  border: 1px solid var(--syn-border);
  padding: var(--syn-space-2);
  border-radius: var(--syn-radius-sm);
  font-size: var(--syn-font-size-sm);
}
</style>
