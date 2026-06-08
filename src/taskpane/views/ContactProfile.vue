<script setup lang="ts">
import { computed, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import ActionButton from '@/taskpane/components/ActionButton.vue'
import Toast from '@/taskpane/components/Toast.vue'
import { getReadItemAsFile, useOutlookItem } from '@/taskpane/composables/useOutlookItem'
import { useSynaplanClient } from '@/taskpane/composables/useSynaplanClient'
import { errorMessage } from '@shared/synaplan-client'
import { renderMarkdown } from '@shared/markdown'
import { setLastRagGroupId } from '@/taskpane/composables/useRoamingSettings'
import { go, selectedContactEmail } from '@/taskpane/router'
import type { RagSearchHit } from '@shared/types'

const { t } = useI18n()
const { call } = useSynaplanClient()
const { item } = useOutlookItem()

const email = computed(() => selectedContactEmail.value ?? '')
const groupKey = computed(() => (email.value ? `contact:${email.value}` : ''))

const query = ref('')
const hits = ref<RagSearchHit[]>([])
const searched = ref(false)
const question = ref('')
const askHistory = ref<{ q: string; a: string }[]>([])
const active = ref<string | null>(null)
const error = ref<string | null>(null)
const status = ref<string | null>(null)
const askChatId = ref<number | undefined>(undefined)

async function run<T>(key: string, fn: () => Promise<T | null>): Promise<T | null> {
  active.value = key
  error.value = null
  try {
    return await fn()
  } catch (err) {
    error.value = errorMessage(err)
    return null
  } finally {
    active.value = null
  }
}

async function search(): Promise<void> {
  // The Synaplan /rag/search endpoint needs a query term — firing it with an
  // empty string (e.g. on view open) just returns an error. Require input.
  const q = query.value.trim()
  if (!groupKey.value || !q) return
  status.value = null
  const r = await run('search', () =>
    call((c) => c.ragSearch({ query: q, groups: [groupKey.value], limit: 15 })),
  )
  searched.value = true
  if (r) hits.value = r
}

async function saveCurrent(): Promise<void> {
  if (!groupKey.value) return
  status.value = null
  await run('save', async () => {
    const file = getReadItemAsFile(item.value)
    const r = await call((c) =>
      c.fileUpload({
        filename: file.filename,
        contentBase64: file.contentBase64,
        mimeType: file.mimeType,
        groupId: groupKey.value,
        processLevel: 'vectorize',
        metadata: {
          contact: email.value,
          from: item.value.from ?? '',
          subject: item.value.subject,
          to: item.value.to.join(', '),
        },
      }),
    )
    if (r) {
      await setLastRagGroupId(groupKey.value).catch(() => undefined)
      // Refresh the list so the freshly saved email shows up, THEN show the
      // confirmation — search() clears `status`, so setting it first would wipe
      // the toast before the user ever sees it.
      await search()
      status.value = t('contactProfile.saved', { group: groupKey.value })
    }
    return r
  })
}

async function ask(): Promise<void> {
  const q = question.value.trim()
  if (!q || !groupKey.value) return
  status.value = null
  // Add the turn up front and stream the answer into it.
  const idx = askHistory.value.push({ q, a: '' }) - 1
  question.value = ''
  const r = await run('ask', async () => {
    // Ground the answer in this contact's saved emails: pull the top RAG
    // snippets for the question and pass them as context to the chat turn.
    const context = await call((c) => c.ragSearch({ query: q, groups: [groupKey.value], limit: 5 }))
    const grounding = (context ?? [])
      .map((h, i) => `[${i + 1}] ${h.filename || 'email'}: ${h.snippet}`)
      .join('\n')
    const augmented = grounding
      ? `About the contact ${email.value}. Use these saved emails as context:\n${grounding}\n\nQuestion: ${q}`
      : `About the contact ${email.value}. Question: ${q}`
    return call((c) =>
      c.chat(
        {
          conversationId: `contact:${email.value}`,
          question: augmented,
          chatId: askChatId.value,
        },
        (textSoFar) => {
          askHistory.value[idx].a = textSoFar
        },
      ),
    )
  })
  if (r) {
    askChatId.value = r.chatId
    askHistory.value[idx].a = r.answer
  } else {
    askHistory.value.splice(idx, 1)
    question.value = q
  }
}
</script>

<template>
  <section class="cp">
    <header class="syn-view-header">
      <button type="button" class="syn-back" @click="go('read')">← {{ t('common.back') }}</button>
      <h2 class="syn-view-title">{{ t('read.contactProfile') }}</h2>
      <p v-if="email" class="syn-view-subtitle">{{ email }}</p>
    </header>

    <p v-if="!email" class="syn-muted">{{ t('contactProfile.noContact') }}</p>

    <template v-else>
      <div class="syn-card">
        <h3 class="syn-card-title">{{ t('contactProfile.results') }}</h3>
        <div class="syn-row">
          <input
            v-model="query"
            type="text"
            class="cp__input"
            :placeholder="t('contactProfile.searchPlaceholder')"
            @keyup.enter="search"
          />
          <ActionButton :block="false" :loading="active === 'search'" @click="search">
            {{ t('contactProfile.search') }}
          </ActionButton>
        </div>

        <ul v-if="hits.length" class="cp__hits">
          <li v-for="(h, i) in hits" :key="i" class="cp__hit">
            <strong>{{ h.filename || t('contactProfile.savedEmail') }}</strong>
            <span class="cp__score">{{ Math.round(h.score * 100) }}%</span>
            <p class="syn-muted">{{ h.snippet }}</p>
          </li>
        </ul>
        <p v-else-if="searched && active !== 'search'" class="syn-muted">
          {{ t('contactProfile.noResults') }}
        </p>

        <ActionButton :loading="active === 'save'" @click="saveCurrent">
          {{ t('contactProfile.saveCurrent') }}
        </ActionButton>
      </div>

      <Toast v-if="status" kind="success" :message="status" />
      <Toast v-if="error" kind="error" :message="error" />

      <div class="syn-card">
        <h3 class="syn-card-title">{{ t('contactProfile.ask') }}</h3>
        <div v-for="(turn, i) in askHistory" :key="i" class="cp__turn">
          <p class="cp__turn-q">{{ turn.q }}</p>
          <!-- AI answer is Markdown/HTML; renderMarkdown sanitises it. -->
          <!-- eslint-disable-next-line vue/no-v-html -->
          <div class="syn-md" v-html="renderMarkdown(turn.a)" />
        </div>
        <div class="syn-row">
          <input
            v-model="question"
            type="text"
            class="cp__input"
            :placeholder="t('contactProfile.askPlaceholder', { email })"
            @keyup.enter="ask"
          />
          <ActionButton :block="false" :loading="active === 'ask'" @click="ask">
            {{ t('read.actions.ask') }}
          </ActionButton>
        </div>
      </div>
    </template>
  </section>
</template>

<style scoped>
.cp {
  padding: var(--syn-space-3);
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-3);
}
/* Color/background/border come from the shared app.css baseline — only the
   row layout (grow to fill the search row) is overridden. */
.cp__input {
  flex: 1;
}
.cp__hits {
  list-style: none;
  padding: 0;
  margin: 0;
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-2);
}
.cp__hit {
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-sm);
  padding: var(--syn-space-2);
}
.cp__score {
  float: right;
  font-size: var(--syn-font-size-sm);
  color: var(--syn-muted);
}
.cp__turn {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-1);
  padding: var(--syn-space-2);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-sm);
}
/* The user's question, set apart from the rendered AI answer below it. */
.cp__turn-q {
  margin: 0;
  font-weight: 600;
  color: var(--syn-muted);
}
</style>
