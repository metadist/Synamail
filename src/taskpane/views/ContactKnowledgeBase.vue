<script setup lang="ts">
import { computed, onMounted, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import ActionButton from '@/taskpane/components/ActionButton.vue'
import Toast from '@/taskpane/components/Toast.vue'
import { getReadItemAsFile, useOutlookItem } from '@/taskpane/composables/useOutlookItem'
import { useSynaplanClient } from '@/taskpane/composables/useSynaplanClient'
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
    error.value = err instanceof Error ? err.message : String(err)
    return null
  } finally {
    active.value = null
  }
}

async function search(): Promise<void> {
  if (!groupKey.value) return
  status.value = null
  const r = await run('search', () =>
    call((c) => c.ragSearch({ query: query.value.trim(), groups: [groupKey.value], limit: 15 })),
  )
  searched.value = true
  if (r) hits.value = r
}

async function saveCurrent(): Promise<void> {
  if (!groupKey.value) return
  status.value = null
  await run('save', async () => {
    const file = await getReadItemAsFile(item.value)
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
      status.value = t('contactKb.saved', { group: groupKey.value })
    }
    return r
  })
}

async function ask(): Promise<void> {
  const q = question.value.trim()
  if (!q || !groupKey.value) return
  status.value = null
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
      c.chat({
        conversationId: `contact:${email.value}`,
        question: augmented,
        chatId: askChatId.value,
      }),
    )
  })
  if (r) {
    askChatId.value = r.chatId
    askHistory.value.push({ q, a: r.answer })
    question.value = ''
  }
}

onMounted(() => {
  if (groupKey.value) void search()
})
</script>

<template>
  <section class="ckb">
    <header class="syn-view-header">
      <button type="button" class="syn-back" @click="go('read')">← {{ t('common.back') }}</button>
      <h2 class="syn-view-title">{{ t('read.contactKb') }}</h2>
      <p v-if="email" class="syn-view-subtitle">{{ email }}</p>
    </header>

    <p v-if="!email" class="syn-muted">{{ t('contactKb.noContact') }}</p>

    <template v-else>
      <div class="syn-card">
        <h3 class="syn-card-title">{{ t('contactKb.results') }}</h3>
        <div class="syn-row">
          <input
            v-model="query"
            type="text"
            class="ckb__input"
            :placeholder="t('contactKb.searchPlaceholder')"
            @keyup.enter="search"
          />
          <ActionButton :block="false" :loading="active === 'search'" @click="search">
            {{ t('contactKb.search') }}
          </ActionButton>
        </div>

        <ul v-if="hits.length" class="ckb__hits">
          <li v-for="(h, i) in hits" :key="i" class="ckb__hit">
            <strong>{{ h.filename || t('contactKb.savedEmail') }}</strong>
            <span class="ckb__score">{{ Math.round(h.score * 100) }}%</span>
            <p class="syn-muted">{{ h.snippet }}</p>
          </li>
        </ul>
        <p v-else-if="searched && active !== 'search'" class="syn-muted">
          {{ t('contactKb.noResults') }}
        </p>

        <ActionButton :loading="active === 'save'" @click="saveCurrent">
          {{ t('contactKb.saveCurrent') }}
        </ActionButton>
      </div>

      <Toast v-if="status" kind="success" :message="status" />
      <Toast v-if="error" kind="error" :message="error" />

      <div class="syn-card">
        <h3 class="syn-card-title">{{ t('contactKb.ask') }}</h3>
        <div v-for="(turn, i) in askHistory" :key="i" class="ckb__turn">
          <p><strong>You:</strong> {{ turn.q }}</p>
          <p><strong>Synaplan:</strong> {{ turn.a }}</p>
        </div>
        <div class="syn-row">
          <input
            v-model="question"
            type="text"
            class="ckb__input"
            :placeholder="t('contactKb.askPlaceholder', { email })"
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
.ckb {
  padding: var(--syn-space-3);
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-3);
}
.ckb__input {
  flex: 1;
  padding: var(--syn-space-2);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-sm);
  font-family: inherit;
}
.ckb__hits {
  list-style: none;
  padding: 0;
  margin: 0;
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-2);
}
.ckb__hit {
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-sm);
  padding: var(--syn-space-2);
}
.ckb__score {
  float: right;
  font-size: var(--syn-font-size-sm);
  color: var(--syn-muted);
}
.ckb__turn {
  padding: var(--syn-space-2);
  border: 1px solid var(--syn-border);
  border-radius: var(--syn-radius-sm);
}
</style>
