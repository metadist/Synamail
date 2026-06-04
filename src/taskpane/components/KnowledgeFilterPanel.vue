<script setup lang="ts">
import { computed, ref } from 'vue'
import { useI18n } from 'vue-i18n'
import ActionButton from '@/taskpane/components/ActionButton.vue'
import Toast from '@/taskpane/components/Toast.vue'
import { useOutlookMailbox } from '@/taskpane/composables/useOutlookMailbox'
import type { MailboxHit } from '@/taskpane/composables/useOutlookMailbox'
import { useSynaplanClient } from '@/taskpane/composables/useSynaplanClient'
import { errorMessage } from '@shared/synaplan-client'
import { getLastRagGroupId, setLastRagGroupId } from '@/taskpane/composables/useRoamingSettings'
import type { RagSearchHit } from '@shared/types'

const emit = defineEmits<{ (e: 'done', group: string): void }>()

const { t } = useI18n()
const { call } = useSynaplanClient()
const { searchMailbox } = useOutlookMailbox()

const query = ref('')
const groupKey = ref(getLastRagGroupId() ?? '')
const searching = ref(false)
const vectorizing = ref(false)
const error = ref<string | null>(null)
const status = ref<string | null>(null)
const searched = ref(false)
const fromOutlook = ref(true)
const mailboxNotice = ref<string | null>(null)

const mailboxHits = ref<MailboxHit[]>([])
const kbHits = ref<RagSearchHit[]>([])
const selected = ref<Set<string>>(new Set())

const effectiveGroup = computed(() => groupKey.value.trim() || slugify(query.value))
const selectedCount = computed(() => selected.value.size)

function slugify(s: string): string {
  return (
    s
      .toLowerCase()
      .trim()
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-+|-+$/g, '')
      .slice(0, 40) || 'topic'
  )
}

function toggle(id: string): void {
  const next = new Set(selected.value)
  if (next.has(id)) next.delete(id)
  else next.add(id)
  selected.value = next
}

async function search(): Promise<void> {
  const q = query.value.trim()
  if (!q) return
  searching.value = true
  error.value = null
  status.value = null
  mailboxNotice.value = null
  const [kbRes, mbRes] = await Promise.allSettled([
    call((c) => c.ragSearch({ query: q, limit: 10 })),
    searchMailbox(q, 10),
  ])

  if (kbRes.status === 'fulfilled') {
    kbHits.value = kbRes.value ?? []
  } else {
    kbHits.value = []
    error.value = errorMessage(kbRes.reason)
  }

  if (mbRes.status === 'fulfilled') {
    mailboxHits.value = mbRes.value.hits
    fromOutlook.value = mbRes.value.fromOutlook
    selected.value = new Set(mbRes.value.hits.map((h) => h.id))
  } else {
    mailboxHits.value = []
    const msg = errorMessage(mbRes.reason)
    mailboxNotice.value = /makeEwsRequest|Elevated permission/i.test(msg)
      ? t('home.search.permissionNeeded')
      : msg
  }

  searched.value = true
  searching.value = false
}

async function vectorize(): Promise<void> {
  const group = effectiveGroup.value
  const picks = mailboxHits.value.filter((h) => selected.value.has(h.id))
  if (picks.length === 0 || !group) return
  vectorizing.value = true
  error.value = null
  status.value = null
  const { getMessageText } = useOutlookMailbox()
  let ok = 0
  try {
    for (const hit of picks) {
      const file = await getMessageText(hit)
      const r = await call((c) =>
        c.fileUpload({
          filename: file.filename,
          contentBase64: file.contentBase64,
          mimeType: file.mimeType,
          groupId: group,
          metadata: { from: hit.from, subject: hit.subject },
          processLevel: 'vectorize',
        }),
      )
      if (r) ok++
    }
    if (ok > 0) await setLastRagGroupId(group)
    status.value = t('home.search.vectorized', { n: ok, group })
    if (ok === picks.length) emit('done', group)
  } catch (err) {
    error.value = errorMessage(err)
  } finally {
    vectorizing.value = false
  }
}
</script>

<template>
  <div class="kf">
    <div class="kf__field">
      <label for="sv-topic">{{ t('home.search.topic') }}</label>
      <input
        id="sv-topic"
        v-model="query"
        type="text"
        :placeholder="t('home.search.topicPlaceholder')"
        @keyup.enter="search"
      />
    </div>

    <div class="kf__field">
      <label for="sv-group">{{ t('home.search.keyword') }}</label>
      <input id="sv-group" v-model="groupKey" type="text" :placeholder="effectiveGroup" />
    </div>

    <ActionButton :loading="searching" :disabled="!query.trim()" @click="search">
      {{ t('home.search.run') }}
    </ActionButton>

    <div v-if="searched" class="kf__results">
      <p class="syn-muted kf__section">
        {{ t('home.search.mailboxResults') }}
        <span v-if="!fromOutlook"> · {{ t('home.search.mockBadge') }}</span>
      </p>
      <p v-if="mailboxNotice" class="kf__notice">{{ mailboxNotice }}</p>
      <p v-else-if="mailboxHits.length === 0" class="syn-muted kf__empty">
        {{ t('home.search.noMailboxHits') }}
      </p>
      <label v-for="h in mailboxHits" :key="h.id" class="kf__hit">
        <input type="checkbox" :checked="selected.has(h.id)" @change="toggle(h.id)" />
        <span class="kf__hit-text">
          <strong>{{ h.subject || '(no subject)' }}</strong>
          <span class="syn-muted">{{ h.from }}</span>
        </span>
      </label>

      <p v-if="kbHits.length" class="syn-muted kf__section">
        {{ t('home.search.kbResults') }}
      </p>
      <div v-for="(h, i) in kbHits" :key="`kb-${i}`" class="kf__hit kf__hit--kb">
        <span class="kf__hit-text">
          <strong>{{ h.filename || '(file)' }}</strong>
          <span class="syn-muted">{{ h.snippet.slice(0, 80) }}</span>
        </span>
      </div>
    </div>

    <Toast v-if="status" kind="success" :message="status" />
    <Toast v-if="error" kind="error" :message="error" />

    <ActionButton
      variant="primary"
      :loading="vectorizing"
      :disabled="selectedCount === 0 || !effectiveGroup"
      @click="vectorize"
    >
      {{ t('home.search.vectorize', { n: selectedCount, group: effectiveGroup }) }}
    </ActionButton>
  </div>
</template>

<style scoped>
.kf {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-3);
}
.kf__field {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-1);
}
.kf__field label {
  font-size: var(--syn-font-size-sm);
  color: var(--syn-muted);
}
.kf__results {
  display: flex;
  flex-direction: column;
  gap: var(--syn-space-1);
}
.kf__section {
  margin: var(--syn-space-2) 0 0;
  font-size: var(--syn-font-size-sm);
}
.kf__empty {
  margin: 0;
  font-size: var(--syn-font-size-sm);
}
.kf__notice {
  margin: 0;
  font-size: var(--syn-font-size-sm);
  color: var(--syn-danger);
}
.kf__hit {
  display: flex;
  gap: var(--syn-space-2);
  align-items: flex-start;
  padding: var(--syn-space-1) 0;
}
.kf__hit-text {
  display: flex;
  flex-direction: column;
  font-size: var(--syn-font-size-sm);
}
.kf__hit--kb {
  padding-left: calc(var(--syn-space-2) + 13px);
}
</style>
